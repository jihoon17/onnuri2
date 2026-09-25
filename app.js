/* =========================================================
   상점가 위치 검색 지도 - 기능 로직
   ========================================================= */

/* -------- 설정 -------- */
// 카카오 개발자센터에서 발급받은 JavaScript 키
const KAKAO_APP_KEY = "01a0f8272692845e09fe8d4c402e6317";

// 엑셀 원본 파일 경로 (이 index.html과 같은 폴더/레포에 올려주세요)
const EXCEL_FILE_URL = "mapData_v1.xlsx";

// 유형별 색상: 전통시장=노랑, 상점가=초록, 골목형상점가=파랑
// base: 구역 기본 색 / dark: 그 구역이 선택(라벨 클릭)됐을 때 진하게 쓰는 색 / marker: 마커 핀 색
const TYPE_COLORS = {
  "전통시장":     { base: "#f5c518", dark: "#8a6d00", marker: "#f5c518" },
  "상점가":       { base: "#3fae5a", dark: "#1f6b34", marker: "#3fae5a" },
  "골목형상점가": { base: "#3b82f6", dark: "#123a63", marker: "#3b82f6" }
};
const DEFAULT_TYPE_COLOR = { base: "#94a3b8", dark: "#475569", marker: "#94a3b8" }; // 미분류 유형 대비

const COLOR_HIGHLIGHT = "#e02424";     // 검색된 특정 주소 (빨간색)

function getTypeColor(type) {
  return TYPE_COLORS[type] || DEFAULT_TYPE_COLOR;
}

/* -------- 전역 상태 -------- */
let map, geocoder, placesService;
let MAP_DATA = { markets: [], parcels: [], zones: [] };

// 세종시 검색 범위 (Places 키워드 검색 시 사용)
// ※ 카카오 radius 최대값은 20000m — 초과 시 검색이 실패할 수 있음
const SEJONG_CENTER_LAT = 36.479934;
const SEJONG_CENTER_LNG = 127.286740;
const SEJONG_SEARCH_RADIUS_M = 20000;

let zoneOverlaysByMarket = {};  // marketName -> [kakao.maps.Polygon, ...] (항상 유지되는 구역 배경)
let marketLabelOverlays = [];   // { marketName, marker, overlay, content } (항상 유지되는 라벨)
let highlightOverlays = [];     // { polygon, marker } (검색 시에만 갱신/삭제되는 강조 표시)
let pickedLocations = []; // [{ id, marker, overlay, latlng, jibunFull, roadFull, areaKey }] (우클릭/꾹 누르기 핀, 여러 개 가능)
let pickedLocationIdSeq = 0;
let pickedViewIndex = -1; // 현재 상세정보에 표시 중인 pickedLocations 인덱스
let pickedAreaPolygons = {}; // areaKey -> { polygon, hitCoords, refCount } (같은 구역 보라색은 하나만)
// 검색 결과도 우클릭과 동일한 상세 패널로 표시하기 위한 상태
let searchDetailItems = []; // [{ latlng, jibunFull, roadFull, parcelAddress, market, parcelId }]
let searchDetailIndex = 0;
let selectedMarket = null;      // 라벨 클릭으로 선택된 상점가 (null이면 선택 없음)
let ignoreClickUntil = 0;       // 롱프레스/라벨 탭 직후 지도 클릭·핀 중복 방지
let lastZoomAt = 0;             // 줌 직후 라벨 탭으로 체크리스트 열리는 것 방지
let multiTouchActive = false;   // 두 손가락(핀치 줌) 중이면 true
let lastMultiTouchAt = 0;       // 멀티터치 종료 시각 (직후 라벨 탭 무시)
let checklistMarket = null;     // 체크리스트가 열려있는 상점가 (null이면 닫힘)
let checklistOverlay = null;    // 체크리스트 흰색 박스(CustomOverlay)
let checklistHighlights = {};   // parcelId -> kakao.maps.Polygon (체크리스트에서 켠 주소 강조)
let activeTypes = new Set(Object.keys(TYPE_COLORS)); // "지도를 볼 기준"에서 켜져있는 유형들

/* =========================================================
   1. 엑셀 원본 읽기
   ========================================================= */
async function loadExcelData(url) {
  const resp = await fetch(url);
  if (!resp.ok) throw new Error("EXCEL_FETCH_FAIL");
  const buf = await resp.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const REQUIRED_SHEETS = ["시장개요", "시장별주소", "구역"];
  const missing = REQUIRED_SHEETS.filter(name => !wb.Sheets[name]);
  if (missing.length) {
    throw new Error(
      `엑셀 파일에 "${missing.join(", ")}" 시트가 없습니다. ` +
      `실제 시트 목록: ${wb.SheetNames.join(", ")}`
    );
  }

  const overviewRows = XLSX.utils.sheet_to_json(wb.Sheets["시장개요"], { defval: "" });
  const addrRows = XLSX.utils.sheet_to_json(wb.Sheets["시장별주소"], { defval: "" });
  const zoneRows = XLSX.utils.sheet_to_json(wb.Sheets["구역"], { defval: "" });

  // 시장개요의 "명칭"은 보통 "지역명 + 구분"으로 되어 있어(예: "보람동 골목형상점가")
  // 시장별주소/구역 시트의 "소속시장"(예: "보람동")과 글자가 다릅니다.
  // 매칭용 baseName은 명칭에서 맨 뒤의 구분(유형) 글자를 떼어낸 값이고,
  // 화면에 보여줄 때는 원래 명칭(name)을 그대로 씁니다.
  function computeBaseName(name, type) {
    const suffix = " " + type;
    if (type && name.endsWith(suffix)) {
      return name.slice(0, -suffix.length).trim();
    }
    return name;
  }

  const markets = overviewRows.map(r => {
    const type = String(r["구분"]).trim();
    const name = String(r["명칭"]).trim();
    return {
      id: r["순번"],
      type,
      name,                              // 화면 표시용 (엑셀 원본 명칭 그대로)
      baseName: computeBaseName(name, type) // 구역/필지 매칭용
    };
  });

  // 경계좌표는 두 가지 형식을 지원:
  // 1) 기존 형식: [{"lat":.., "lng":..}, ...]
  // 2) GeoJSON 형식: {"type":"Polygon"|"MultiPolygon", "coordinates": [...]}
  //    (GeoJSON은 좌표 순서가 [경도, 위도]이고, MultiPolygon은 폴리곤이 여러 겹 중첩된 배열입니다.
  //     이 데이터는 폴리곤/링이 항상 1개씩이라 첫 번째 폴리곤의 바깥 링만 사용합니다.)
  function geoJsonToFlatPoints(geo, sheetLabel, rowNo) {
    let ring;
    if (geo.type === "Polygon") {
      ring = geo.coordinates[0];
    } else if (geo.type === "MultiPolygon") {
      ring = geo.coordinates[0][0];
    } else {
      throw new Error(`"${sheetLabel}" 시트 ${rowNo}번째 행: 지원하지 않는 GeoJSON 타입(${geo.type})입니다.`);
    }
    if (!ring || !ring.length) {
      throw new Error(`"${sheetLabel}" 시트 ${rowNo}번째 행의 좌표가 비어 있습니다.`);
    }
    return ring.map(pos => ({ lat: pos[1], lng: pos[0] }));
  }

  function parseCoords(raw, sheetLabel, rowNo) {
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (e) {
      throw new Error(`"${sheetLabel}" 시트 ${rowNo}번째 행의 경계좌표 형식이 올바르지 않습니다.`);
    }

    if (Array.isArray(parsed)) {
      return parsed; // 기존 [{lat,lng}, ...] 형식
    }
    if (parsed && typeof parsed === "object" && parsed.type && parsed.coordinates) {
      return geoJsonToFlatPoints(parsed, sheetLabel, rowNo); // GeoJSON 형식
    }
    throw new Error(`"${sheetLabel}" 시트 ${rowNo}번째 행의 경계좌표 형식을 인식할 수 없습니다.`);
  }

  const parcels = addrRows.map((r, i) => {
    // 지번주소: 기존 "주소" 컬럼 (또는 "지번주소")
    const jibun = String(r["주소"] ?? r["지번주소"] ?? "").trim();
    // 도로명주소: 새로 추가된 컬럼
    const road = String(r["도로명주소"] ?? "").trim();
    return {
      id: r["순번"],
      market: String(r["소속시장"]).trim(),
      address: jibun,           // 지번주소 (기존 호환)
      roadAddress: road,        // 도로명주소
      coords: parseCoords(r["경계좌표"], "시장별주소", i + 2)
    };
  });

  const zones = zoneRows.map((r, i) => ({
    id: r["순번"],
    market: String(r["소속시장"]).trim(),
    zoneNo: r["구역순번"],
    coords: parseCoords(r["경계좌표"], "구역", i + 2)
  }));

  return { markets, parcels, zones };
}

/* =========================================================
   2. 카카오맵 SDK 로드 / 지도 초기화
   ========================================================= */
function loadKakaoSdk(appkey) {
  return new Promise((resolve, reject) => {
    if (!appkey || appkey.startsWith("여기에")) {
      reject(new Error("NO_KEY"));
      return;
    }
    const script = document.createElement("script");
    script.src = `https://dapi.kakao.com/v2/maps/sdk.js?appkey=${appkey}&autoload=false&libraries=services`;
    script.onload = () => {
      kakao.maps.load(() => resolve());
    };
    script.onerror = () => reject(new Error("SDK_LOAD_FAIL"));
    document.head.appendChild(script);
  });
}

function initMap() {
  const container = document.getElementById("map");
  // 초기 중심: 위도 36.479934, 경도 127.286740 / 줌 레벨 4
  const defaultCenter = new kakao.maps.LatLng(36.479934, 127.286740);
  map = new kakao.maps.Map(container, {
    center: defaultCenter,
    level: 4,
    disableDoubleClickZoom: true
  });
  // PC: 카카오 기본 드래그 사용 / 터치 기기: 네이버맵식 직접 제스처(아래) 사용
  const isTouchDevice = ("ontouchstart" in window) || (navigator.maxTouchPoints > 0);
  if (isTouchDevice) {
    // 카카오 기본 터치가 패드에서 동작하지 않는 경우가 많아 직접 제어
    map.setDraggable(false);
  } else {
    map.setDraggable(true);
  }
  map.setZoomable(true); // +/- 버튼 줌은 유지
  geocoder = new kakao.maps.services.Geocoder();
  placesService = new kakao.maps.services.Places();

  // 지도 / 스카이뷰 전환 버튼 (왼쪽 상단) — 회색 부가 글자 없이 기본만
  const mapTypeControl = new kakao.maps.MapTypeControl();
  map.addControl(mapTypeControl, kakao.maps.ControlPosition.TOPLEFT);

  // 카카오 기본 ZoomControl 은 쓰지 않음 (회색 글자·크기 문제) → 커스텀 #mapZoom 사용
  const zoomInBtn = document.getElementById("zoomInBtn");
  const zoomOutBtn = document.getElementById("zoomOutBtn");
  if (zoomInBtn) {
    zoomInBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (!map) return;
      try { map.setLevel(Math.max(1, map.getLevel() - 1), { animate: false }); } catch (_) { map.setLevel(Math.max(1, map.getLevel() - 1)); }
    });
  }
  if (zoomOutBtn) {
    zoomOutBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (!map) return;
      try { map.setLevel(Math.min(14, map.getLevel() + 1), { animate: false }); } catch (_) { map.setLevel(Math.min(14, map.getLevel() + 1)); }
    });
  }

  // 줌 레벨 변경 시 라벨 표시/숨김
  kakao.maps.event.addListener(map, "zoom_changed", () => {
    lastZoomAt = Date.now();
    updateLabelVisibilityByZoom();
  });

  function containerPointToLatLng(clientX, clientY) {
    const rect = container.getBoundingClientRect();
    const x = clientX - rect.left;
    const y = clientY - rect.top;
    return map.getProjection().coordsFromContainerPoint(new kakao.maps.Point(x, y));
  }

  // ---------- 제스처 상태 ----------
  let longPressTimer = null;
  let longPressStartXY = null;
  let lastTouchAt = 0;
  let lastPickAt = 0;
  let lastPickKey = "";

  // 한 손가락 팬 (직전 좌표 대비 증분 — 미끄러짐 제거)
  let panActive = false;
  let panLastTouch = null; // {x, y} 직전 터치 위치
  let touchPanPendingDx = 0;
  let touchPanPendingDy = 0;
  let touchPanRaf = null;

  // 두 손가락 핀치 줌 — 마우스 휠처럼 "한 단계씩"만 적용 (지도·구역이 같이 붙도록)
  let pinchActive = false;
  let pinchStartDist = 0;
  // 한 단계 줌에 필요한 거리 비율 (벌리면 확대, 모으면 축소)
  const PINCH_IN_RATIO = 1.22;   // 22% 이상 벌리면 1레벨 확대
  const PINCH_OUT_RATIO = 0.82;  // 18% 이상 모으면 1레벨 축소

  const LONG_PRESS_MS = 550;
  const MOVE_CANCEL_PX = 10;
  const PICK_DEDUP_MS = 900;
  const MIN_LEVEL = 1;
  const MAX_LEVEL = 14;

  function getTouchPanSensitivity() {
    const w = window.innerWidth || 400;
    if (w >= 600) return 1.0;
    return 1.0;
  }

  function flushTouchPan() {
    touchPanRaf = null;
    if (!map || !panActive) return;
    const dx = touchPanPendingDx;
    const dy = touchPanPendingDy;
    touchPanPendingDx = 0;
    touchPanPendingDy = 0;
    if (Math.abs(dx) < 0.01 && Math.abs(dy) < 0.01) return;
    try {
      const proj = map.getProjection();
      const center = map.getCenter();
      const pt = proj.containerPointFromCoords(center);
      const desired = new kakao.maps.Point(pt.x - dx, pt.y - dy);
      map.setCenter(proj.coordsFromContainerPoint(desired));
    } catch (_) {}
  }

  function scheduleTouchPan() {
    if (touchPanRaf == null) touchPanRaf = requestAnimationFrame(flushTouchPan);
  }

  /** 마우스 휠처럼 레벨을 1단계만 바꾸고, 기준 거리를 리셋해 연속 점프를 막음 */
  function applyDiscretePinchStep(dir, currentDist) {
    if (!map) return;
    const cur = map.getLevel();
    const next = Math.max(MIN_LEVEL, Math.min(MAX_LEVEL, cur + dir));
    if (next === cur) return;
    try {
      map.setLevel(next, { animate: false });
    } catch (_) {
      map.setLevel(next);
    }
    lastZoomAt = Date.now();
    // 다음 1단계를 위해 기준 거리 재설정 (휠 한 칸과 동일)
    pinchStartDist = currentDist || pinchStartDist;
  }

  const isInteractiveTarget = (target) => {
    if (!target || !target.closest) return false;
    return !!target.closest(
      ".market-label, button, a, input, label, .suggest-item, .addr-checklist-panel, .type-chip, .mic-btn, .map-zoom, .map-joystick, .panel-toggle"
    );
  };

  const clearLongPress = () => {
    if (longPressTimer !== null) {
      clearTimeout(longPressTimer);
      longPressTimer = null;
    }
    longPressStartXY = null;
  };

  const safeMapPick = (latlng) => {
    if (!latlng) return;
    const key = `${latlng.getLat().toFixed(5)},${latlng.getLng().toFixed(5)}`;
    const now = Date.now();
    if (key === lastPickKey && now - lastPickAt < PICK_DEDUP_MS) return;
    lastPickKey = key;
    lastPickAt = now;
    ignoreClickUntil = now + 700;
    handleMapPick(latlng);
  };

  const touchDist = (a, b) =>
    Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);

  const startLongPress = (clientX, clientY) => {
    clearLongPress();
    longPressStartXY = { x: clientX, y: clientY };
    longPressTimer = setTimeout(() => {
      longPressTimer = null;
      longPressStartXY = null;
      if (panActive || pinchActive || multiTouchActive) return;
      safeMapPick(containerPointToLatLng(clientX, clientY));
    }, LONG_PRESS_MS);
  };

  // 우클릭(마우스)
  container.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    if (Date.now() - lastTouchAt < 1200) return;
    safeMapPick(containerPointToLatLng(e.clientX, e.clientY));
  });

  kakao.maps.event.addListener(map, "click", (mouseEvent) => {
    if (Date.now() < ignoreClickUntil) return;
    handleMapClick(mouseEvent.latLng);
  });

  container.addEventListener("click", (e) => {
    if (Date.now() < ignoreClickUntil) return;
    if (e.button !== 0) return;
    if (isInteractiveTarget(e.target)) return;
    handleMapClick(containerPointToLatLng(e.clientX, e.clientY));
  });

  // ===== 터치 제스처: 1손가락 이동(증분) / 2손가락 핀치(rAF) =====
  container.addEventListener("touchstart", (e) => {
    lastTouchAt = Date.now();

    if (isInteractiveTarget(e.target)) {
      clearLongPress();
      panActive = false;
      pinchActive = false;
      panLastTouch = null;
      return;
    }

    if (e.touches.length >= 2) {
      multiTouchActive = true;
      lastMultiTouchAt = Date.now();
      clearLongPress();
      panActive = false;
      panLastTouch = null;
      pinchActive = true;
      pinchStartDist = touchDist(e.touches[0], e.touches[1]) || 1;
      return;
    }

    if (e.touches.length === 1) {
      multiTouchActive = false;
      pinchActive = false;
      const t = e.touches[0];
      panActive = false;
      panLastTouch = { x: t.clientX, y: t.clientY };
      startLongPress(t.clientX, t.clientY);
    }
  }, { passive: true });

  container.addEventListener("touchmove", (e) => {
    // ----- 두 손가락: 핀치 줌 (마우스 휠처럼 임계값마다 1레벨만) -----
    if (e.touches.length >= 2) {
      multiTouchActive = true;
      lastMultiTouchAt = Date.now();
      clearLongPress();
      panActive = false;
      panLastTouch = null;

      if (!pinchActive) {
        pinchActive = true;
        pinchStartDist = touchDist(e.touches[0], e.touches[1]) || 1;
      }

      e.preventDefault();
      if (!map) return;
      const dist = touchDist(e.touches[0], e.touches[1]) || pinchStartDist;
      const ratio = dist / (pinchStartDist || 1);
      // 카카오: level 작을수록 확대 → 벌리면 -1, 모으면 +1
      if (ratio >= PINCH_IN_RATIO) {
        applyDiscretePinchStep(-1, dist);
      } else if (ratio <= PINCH_OUT_RATIO) {
        applyDiscretePinchStep(1, dist);
      }
      return;
    }

    // ----- 한 손가락: 증분 이동 (미끄러짐 없음) -----
    if (e.touches.length === 1 && panLastTouch) {
      const t = e.touches[0];
      const rawDx = t.clientX - panLastTouch.x;
      const rawDy = t.clientY - panLastTouch.y;
      const dist = Math.hypot(rawDx, rawDy);

      if (dist > MOVE_CANCEL_PX || panActive) {
        clearLongPress();
        panActive = true;
        e.preventDefault();
        const sens = getTouchPanSensitivity();
        const dx = rawDx * sens;
        const dy = rawDy * sens;
        panLastTouch = { x: t.clientX, y: t.clientY };
        touchPanPendingDx += dx;
        touchPanPendingDy += dy;
        scheduleTouchPan();
      }
    }
  }, { passive: false });

  container.addEventListener("touchend", (e) => {
    clearLongPress();
    const remaining = e.touches ? e.touches.length : 0;

    if (remaining >= 2) {
      multiTouchActive = true;
      return;
    }

    if (remaining === 1) {
      multiTouchActive = false;
      lastMultiTouchAt = Date.now();
      pinchActive = false;
      const t = e.touches[0];
      panLastTouch = { x: t.clientX, y: t.clientY };
      panActive = false;
      return;
    }

    if (multiTouchActive || pinchActive) lastMultiTouchAt = Date.now();
    multiTouchActive = false;
    pinchActive = false;
    panActive = false;
    panLastTouch = null;
    touchPanPendingDx = 0;
    touchPanPendingDy = 0;
  }, { passive: true });

  container.addEventListener("touchcancel", () => {
    clearLongPress();
    multiTouchActive = false;
    pinchActive = false;
    panActive = false;
    panLastTouch = null;
    touchPanPendingDx = 0;
    touchPanPendingDy = 0;
    lastMultiTouchAt = Date.now();
  }, { passive: true });

  // 마우스 길게 누르기 (PC / 패드+마우스)
  container.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    if (Date.now() - lastTouchAt < 800) return;
    if (isInteractiveTarget(e.target)) return;
    startLongPress(e.clientX, e.clientY);
  });
  container.addEventListener("mousemove", (e) => {
    if (!longPressStartXY) return;
    const moved = Math.abs(e.clientX - longPressStartXY.x) + Math.abs(e.clientY - longPressStartXY.y);
    if (moved > MOVE_CANCEL_PX) clearLongPress();
  });
  container.addEventListener("mouseup", () => clearLongPress());
  container.addEventListener("mouseleave", () => clearLongPress());
}

function showFallback(message) {
  const wrap = document.querySelector(".map-wrap");
  const fallback = document.createElement("div");
  fallback.className = "map-fallback";
  fallback.innerHTML = `
    <div>🗺️ ${message}</div>
    <div>app.js 상단의 <code>KAKAO_APP_KEY</code> 값을<br>
    카카오 개발자센터에서 발급받은 JavaScript 키로 교체해주세요.</div>
  `;
  wrap.appendChild(fallback);
}

/* =========================================================
   3. 데이터 헬퍼
   ========================================================= */
function toLatLngPath(coords) {
  return coords.map(c => new kakao.maps.LatLng(c.lat, c.lng));
}

function getZonesForMarket(marketName) {
  return MAP_DATA.zones.filter(z => z.market === marketName);
}

function getMarketType(marketName) {
  const m = MAP_DATA.markets.find(m => m.baseName === marketName);
  return m ? m.type : null;
}

// 매칭용 baseName으로 화면에 보여줄 원래 명칭을 찾음 (없으면 baseName 그대로 반환)
function getMarketDisplayName(marketName) {
  const m = MAP_DATA.markets.find(m => m.baseName === marketName);
  return m ? m.name : marketName;
}

// 라벨용 짧은 이름: 골목형상점가는 "도담동(골목형)" 형태로 표시
function getMarketLabelText(market) {
  if (market.type === "골목형상점가") {
    return `${market.baseName}(골목형)`;
  }
  return market.name;
}

// 도로명주소에서 "세종시" / "세종특별자치시" 등 시 이름 제거한 축약형
function abbreviateRoadAddress(road) {
  if (!road) return "";
  return road
    .replace(/^세종특별자치시\s*/, "")
    .replace(/^세종시\s*/, "")
    .trim();
}

// 체크리스트 등에 표시할 주소 문자열: 지번주소(도로명주소)
function formatParcelAddress(parcel) {
  const jibun = parcel.address || "";
  const roadShort = abbreviateRoadAddress(parcel.roadAddress || "");
  if (jibun && roadShort) return `${jibun}(${roadShort})`;
  if (jibun) return jibun;
  if (roadShort) return roadShort;
  return "";
}

// 필지 좌표의 대략적인 중심점
function getParcelCenter(parcel) {
  if (!parcel || !parcel.coords || !parcel.coords.length) return null;
  let lat = 0, lng = 0;
  parcel.coords.forEach(c => { lat += c.lat; lng += c.lng; });
  const n = parcel.coords.length;
  return new kakao.maps.LatLng(lat / n, lng / n);
}

// 구역 안내 박스 유형별 색상 (전통시장=연노랑, 상점가=연초록, 골목형=연파랑)
const ZONE_NOTICE_THEME = {
  "전통시장": {
    bg: "#fff8d6", border: "#e6b800", text: "#6b5500", strong: "#4a3a00", link: "#5c4a00"
  },
  "상점가": {
    bg: "#e5f6ea", border: "#3fae5a", text: "#1f6b34", strong: "#145228", link: "#0f4d1f"
  },
  "골목형상점가": {
    bg: "#dbeafe", border: "#3b82f6", text: "#1e40af", strong: "#1e3a8a", link: "#1d4ed8"
  },
  out: {
    bg: "#f8e8e8", border: "#dc3545", text: "#721c24", strong: "#491217", link: "#491217"
  }
};

function getZoneNoticeTheme(marketName) {
  if (!marketName) return ZONE_NOTICE_THEME.out;
  const type = getMarketType(marketName);
  return ZONE_NOTICE_THEME[type] || ZONE_NOTICE_THEME.out;
}

/** 검색·우클릭 공통: 구역 안내 HTML (유형별 색 + 다음 줄 구역명 링크) */
function buildZoneNoticeHtml(marketNames) {
  const names = (marketNames || []).filter(Boolean);
  if (!names.length) {
    const t = ZONE_NOTICE_THEME.out;
    return `
      <div class="result-zone-notice result-zone-notice--typed" style="--zn-bg:${t.bg};--zn-border:${t.border};--zn-text:${t.text};--zn-strong:${t.strong};--zn-link:${t.link}">
        <div class="zone-notice-label">📍 검색 주소가 포함된 구역:</div>
        <div class="zone-notice-name">
          <strong>해당 없음</strong>
          <span class="zone-out-sub">(등록된 상점가 구역 밖)</span>
        </div>
      </div>`;
  }
  // 여러 구역이면 첫 번째 유형 색을 기준으로 사용
  const theme = getZoneNoticeTheme(names[0]);
  const links = names.map(m =>
    `<a href="#" class="zone-link" data-market="${m}">${getMarketDisplayName(m)}</a>`
  ).join('<span class="zone-sep">, </span>');
  return `
    <div class="result-zone-notice result-zone-notice--typed" style="--zn-bg:${theme.bg};--zn-border:${theme.border};--zn-text:${theme.text};--zn-strong:${theme.strong};--zn-link:${theme.link}">
      <div class="zone-notice-label">📍 검색 주소가 포함된 구역:</div>
      <div class="zone-notice-name">${links}</div>
    </div>`;
}

/** 검색·우클릭 공통: 상세 정보 행 (상호명 선택 / 지번/도로명/위도/경도) */
function buildDetailRowsHtml(latlng, jibunFull, roadFull, placeName) {
  const lat = latlng && typeof latlng.getLat === "function" ? latlng.getLat().toFixed(6) : "확인되지 않음";
  const lng = latlng && typeof latlng.getLng === "function" ? latlng.getLng().toFixed(6) : "확인되지 않음";
  const placeRow = placeName
    ? `<div class="detail-row"><span class="d-label">상호명</span><span class="d-value">${placeName}</span></div>`
    : "";
  return `
    <div class="detail-block">
      ${placeRow}
      <div class="detail-row"><span class="d-label">지번주소</span><span class="d-value">${jibunFull || "확인되지 않음"}</span></div>
      <div class="detail-row"><span class="d-label">도로명주소</span><span class="d-value">${roadFull || "확인되지 않음"}</span></div>
      <div class="detail-row"><span class="d-label">위도</span><span class="d-value">${lat}</span></div>
      <div class="detail-row"><span class="d-label">경도</span><span class="d-value">${lng}</span></div>
    </div>`;
}

function bindZoneLinkClicks(container) {
  if (!container) return;
  container.querySelectorAll(".zone-link").forEach((el) => {
    el.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const m = el.getAttribute("data-market");
      if (m) focusMarketZone(m);
    });
  });
}

// 좌표 목록 중 위도(lat)가 가장 높은 좌표 반환
function getHighestLatPoint(coordsList) {
  let best = null;
  coordsList.forEach(c => {
    if (!best || c.lat > best.lat) best = c;
  });
  return best;
}

// 검색어와 매칭되는 필지(주소) 목록 (시장명 또는 지번/도로명 주소 포함검색)
function searchParcels(query) {
  const q = query.trim();
  if (!q) return [];
  return MAP_DATA.parcels.filter(p =>
    p.market.includes(q) ||
    p.address.includes(q) ||
    (p.roadAddress && p.roadAddress.includes(q)) ||
    q.includes(p.market)
  );
}

// 검색어와 정확히(또는 강하게) 일치하는 필지 우선 매칭
function searchParcelsExact(query) {
  const q = query.trim();
  if (!q) return [];
  const exact = MAP_DATA.parcels.filter(p =>
    p.address === q ||
    (p.roadAddress && p.roadAddress === q) ||
    abbreviateRoadAddress(p.roadAddress || "") === q
  );
  if (exact.length) return exact;
  // 지번/도로명이 검색어로 끝나거나 포함되는 경우
  return MAP_DATA.parcels.filter(p =>
    p.address.includes(q) ||
    (p.roadAddress && p.roadAddress.includes(q))
  );
}

function fitBoundsToPaths(paths) {
  if (!paths.length) return;
  const bounds = new kakao.maps.LatLngBounds();
  paths.forEach(path => path.forEach(ll => bounds.extend(ll)));
  map.setBounds(bounds);
}

/* =========================================================
   4. 구역(배경) 레이어 - 항상 모든 구역을 파란색 계열로 표시
   ========================================================= */
function drawAllZonesBase() {
  Object.values(zoneOverlaysByMarket).flat().forEach(p => p.setMap(null));
  zoneOverlaysByMarket = {};

  MAP_DATA.markets.forEach(m => {
    const zones = getZonesForMarket(m.baseName);
    const colors = getTypeColor(m.type);
    const visible = activeTypes.has(m.type) ? map : null;
    const polygons = zones.map(z => {
      const path = toLatLngPath(z.coords);
      return new kakao.maps.Polygon({
        map: visible,
        path,
        strokeWeight: 2,
        strokeColor: colors.base,
        strokeOpacity: 0.9,
        fillColor: colors.base,
        fillOpacity: 0.4,
        zIndex: 1
      });
    });
    zoneOverlaysByMarket[m.baseName] = polygons;
  });
}

// 현재 selectedMarket 상태에 맞춰 구역 색상을 다시 칠함 (다시 그리지 않고 옵션만 변경)
// 기본 색은 상점가 유형(전통시장/상점가/골목형상점가)에 따라 다르며, 선택된 구역만 그
// 유형의 진한 색으로 강조. 선택 안 된 나머지 구역은 옅어지지 않고 자기 유형의 기본색을 유지.
function applyZoneColorState() {
  Object.entries(zoneOverlaysByMarket).forEach(([marketName, polygons]) => {
    const type = getMarketType(marketName);
    const colors = getTypeColor(type);
    let color, fillOpacity;
    if (marketName === selectedMarket) {
      color = colors.dark;
      fillOpacity = 0.6;
    } else {
      color = colors.base;
      fillOpacity = 0.4;
    }
    polygons.forEach(p => p.setOptions({ strokeColor: color, fillColor: color, fillOpacity }));
  });

  marketLabelOverlays.forEach(o => {
    const isSelected = o.marketName === selectedMarket;
    const colors = getTypeColor(o.type);
    o.content.style.setProperty("--sel-color", colors.dark);
    o.content.classList.toggle("selected", isSelected);
  });
}

// 라벨 클릭 토글:
// 1번 클릭: 검색 빨간 강조 제거 + 구역 진한 색 강조 + 주소 체크리스트 표시
// 같은 라벨 다시 클릭: 구역 강조 해제 + 체크리스트 닫기 (+ 검색 빨간 강조도 제거)
// 체크리스트의 X 버튼을 누르면 강조는 유지한 채 체크리스트만 닫힘
// 지도의 줌 레벨/범위는 변경하지 않고 색상/체크리스트만 바꿈
function selectMarketByLabel(marketName) {
  // 검색으로 생긴 빨간 강조는 라벨 클릭 시 항상 제거
  clearHighlights();

  if (selectedMarket === marketName) {
    // 이미 강조된 라벨을 다시 클릭 → 강조 해제 + 체크리스트 닫기
    closeChecklist();
    selectedMarket = null;
    applyZoneColorState();
    return;
  }

  closeChecklist();
  selectedMarket = marketName;
  applyZoneColorState();
  openChecklist(marketName);
}

// 체크리스트에서 켠 주소 강조 폴리곤을 모두 지움
function clearChecklistHighlights() {
  Object.values(checklistHighlights).forEach(polygon => polygon.setMap(null));
  checklistHighlights = {};
}

// 체크리스트(흰색 박스)와 그 안에서 켠 주소 강조를 모두 닫음 (X 버튼 클릭 시에도 호출됨)
function closeChecklist() {
  if (checklistOverlay) {
    // 예전 CustomOverlay 호환
    try { checklistOverlay.setMap(null); } catch (e) {}
    checklistOverlay = null;
  }
  const floating = document.getElementById("addrChecklistPanel");
  if (floating) floating.remove();
  clearChecklistHighlights();
  checklistMarket = null;
}

// 체크리스트에서 체크박스를 켰을 때 그 주소(필지) 위치를 구역 안에 강조 표시
function drawChecklistHighlight(parcel) {
  const path = toLatLngPath(parcel.coords);
  return new kakao.maps.Polygon({
    map,
    path,
    strokeWeight: 3,
    strokeColor: COLOR_HIGHLIGHT,
    strokeOpacity: 1,
    fillColor: COLOR_HIGHLIGHT,
    fillOpacity: 0.55,
    zIndex: 15
  });
}

// 특정 상점가에 속한 주소 체크리스트(흰색 박스)
// - 지도 위 플로팅 패널로 표시 (드래그 / 리사이즈 가능)
// - 제목 바를 좌클릭 드래그하면 이동, 오른쪽 하단 핸들로 크기 조절
function openChecklist(marketName) {
  closeChecklist();

  const parcels = MAP_DATA.parcels.filter(p => p.market === marketName);
  const zones = getZonesForMarket(marketName);
  const allCoords = zones.flatMap(z => z.coords);
  const topPoint = getHighestLatPoint(allCoords);

  const mapWrap = document.querySelector(".map-wrap") || document.getElementById("map")?.parentElement;
  if (!mapWrap) return;

  // 체크리스트는 지도 우측 상단에 표시 (기본 크기의 4/5)
  const wrapRect = mapWrap.getBoundingClientRect();
  const panelW = 208;
  const initLeft = Math.max(8, wrapRect.width - panelW - 16);
  const initTop = 16;

  const panel = document.createElement("div");
  panel.id = "addrChecklistPanel";
  panel.className = "addr-checklist-panel";
  panel.style.left = initLeft + "px";
  panel.style.top = initTop + "px";

  // 헤더(드래그 핸들)
  const header = document.createElement("div");
  header.className = "addr-checklist-header";

  const title = document.createElement("div");
  title.className = "addr-checklist-title";
  title.textContent = `${getMarketDisplayName(marketName)} 주소`;
  header.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "addr-checklist-close";
  closeBtn.setAttribute("aria-label", "체크리스트 닫기");
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    closeChecklist();
  });
  header.appendChild(closeBtn);
  panel.appendChild(header);

  const body = document.createElement("div");
  body.className = "addr-checklist-body";

  if (!parcels.length) {
    const empty = document.createElement("div");
    empty.className = "addr-checklist-empty";
    empty.textContent = "등록된 주소가 없습니다.";
    body.appendChild(empty);
  }

  parcels.forEach(p => {
    const row = document.createElement("label");
    row.className = "addr-checklist-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        checklistHighlights[p.id] = drawChecklistHighlight(p);
      } else if (checklistHighlights[p.id]) {
        checklistHighlights[p.id].setMap(null);
        delete checklistHighlights[p.id];
      }
    });

    const span = document.createElement("span");
    span.textContent = formatParcelAddress(p);

    row.appendChild(checkbox);
    row.appendChild(span);
    body.appendChild(row);
  });
  panel.appendChild(body);

  // 리사이즈 핸들
  const resizeHandle = document.createElement("div");
  resizeHandle.className = "addr-checklist-resize";
  resizeHandle.title = "드래그해서 크기 조절";
  panel.appendChild(resizeHandle);

  mapWrap.appendChild(panel);
  checklistMarket = marketName;
  // CustomOverlay 자리는 쓰지 않음 (호환용 null)
  checklistOverlay = null;

  // ---- 드래그 (헤더) ----
  let dragging = false;
  let dragOffsetX = 0;
  let dragOffsetY = 0;

  header.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    if (e.target === closeBtn || closeBtn.contains(e.target)) return;
    dragging = true;
    const rect = panel.getBoundingClientRect();
    const wrapRect = mapWrap.getBoundingClientRect();
    dragOffsetX = e.clientX - rect.left;
    dragOffsetY = e.clientY - rect.top;
    panel.classList.add("dragging");
    e.preventDefault();
  });

  // ---- 리사이즈 ----
  let resizing = false;
  let startW = 0, startH = 0, startX = 0, startY = 0;

  resizeHandle.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    resizing = true;
    startW = panel.offsetWidth;
    startH = panel.offsetHeight;
    startX = e.clientX;
    startY = e.clientY;
    panel.classList.add("resizing");
    e.preventDefault();
    e.stopPropagation();
  });

  const onMove = (e) => {
    if (dragging) {
      const wrapRect = mapWrap.getBoundingClientRect();
      let left = e.clientX - wrapRect.left - dragOffsetX;
      let top = e.clientY - wrapRect.top - dragOffsetY;
      left = Math.max(0, Math.min(left, wrapRect.width - 80));
      top = Math.max(0, Math.min(top, wrapRect.height - 40));
      panel.style.left = left + "px";
      panel.style.top = top + "px";
    } else if (resizing) {
      const dw = e.clientX - startX;
      const dh = e.clientY - startY;
      const newW = Math.max(144, Math.min(336, startW + dw));
      const newH = Math.max(96, Math.min(384, startH + dh));
      panel.style.width = newW + "px";
      panel.style.height = newH + "px";
    }
  };

  const onUp = () => {
    if (dragging || resizing) {
      dragging = false;
      resizing = false;
      panel.classList.remove("dragging", "resizing");
    }
  };

  document.addEventListener("mousemove", onMove);
  document.addEventListener("mouseup", onUp);

  // 패널이 제거될 때 리스너 정리
  const observer = new MutationObserver(() => {
    if (!document.getElementById("addrChecklistPanel")) {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      observer.disconnect();
    }
  });
  observer.observe(mapWrap, { childList: true });
}

// 유형 색상에 맞는 핀 모양 마커 이미지 생성 (SVG data URL)
function createMarkerImage(color) {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="38" viewBox="0 0 28 38">` +
    `<path d="M14 0C6.3 0 0 6.3 0 14c0 10.5 14 24 14 24s14-13.5 14-24C28 6.3 21.7 0 14 0z" fill="${color}" stroke="#ffffff" stroke-width="1.5"/>` +
    `<circle cx="14" cy="14" r="5.5" fill="#ffffff"/>` +
    `</svg>`;
  const url = "data:image/svg+xml;charset=UTF-8," + encodeURIComponent(svg);
  return new kakao.maps.MarkerImage(
    url,
    new kakao.maps.Size(28, 38),
    { offset: new kakao.maps.Point(14, 38) }
  );
}

/* =========================================================
   5. 상점가 명칭 라벨 레이어 - 각 구역 중 위도가 가장 높은 지점에 표시
   ========================================================= */
// 라벨이 보이는 최대 줌 레벨 (이보다 숫자가 크면 = 더 줌아웃되면 라벨 숨김, 마커는 유지)
// 카카오맵: level 숫자가 클수록 더 멀리 보임
const LABEL_MAX_LEVEL = 6;

function updateLabelVisibilityByZoom() {
  if (!map) return;
  const level = map.getLevel();
  const showLabel = level <= LABEL_MAX_LEVEL;

  marketLabelOverlays.forEach(o => {
    const typeVisible = activeTypes.has(o.type);
    // 마커는 유형 필터만 따름
    o.marker.setMap(typeVisible ? map : null);
    // 라벨은 유형 필터 + 줌 레벨 둘 다 만족해야 표시
    o.overlay.setMap(typeVisible && showLabel ? map : null);
  });
}

function drawMarketLabels() {
  marketLabelOverlays.forEach(o => {
    if (o.marker) o.marker.setMap(null);
    if (o.overlay) o.overlay.setMap(null);
  });
  marketLabelOverlays = [];

  MAP_DATA.markets.forEach(m => {
    const zones = getZonesForMarket(m.baseName);
    if (!zones.length) return;

    const allCoords = zones.flatMap(z => z.coords);
    const topPoint = getHighestLatPoint(allCoords);
    if (!topPoint) return;

    const position = new kakao.maps.LatLng(topPoint.lat, topPoint.lng);
    const colors = getTypeColor(m.type);
    const typeVisible = activeTypes.has(m.type);
    const level = map ? map.getLevel() : 5;
    const showLabel = level <= LABEL_MAX_LEVEL;
    const visible = typeVisible ? map : null;
    const labelVisible = typeVisible && showLabel ? map : null;

    const marker = new kakao.maps.Marker({
      map: visible,
      position,
      image: createMarkerImage(colors.marker),
      zIndex: 20
    });

    const content = document.createElement("div");
    content.className = "market-label";
    content.textContent = getMarketLabelText(m);
    content.style.setProperty("--sel-color", colors.dark);
    content.style.cursor = "pointer";
    content.style.touchAction = "manipulation";
    // 터치·마우스 모두 동작 (짧은 시간에 중복 호출 방지)
    let lastActivateAt = 0;
    const onLabelActivate = (e) => {
      if (e) {
        e.preventDefault();
        e.stopPropagation();
      }
      const now = Date.now();
      // 두 손가락 제스처 중·직후, 또는 줌 직후에는 체크리스트를 열지 않음
      if (multiTouchActive) return;
      if (now - lastMultiTouchAt < 300) return;
      if (now - lastZoomAt < 300) return;
      if (now - lastActivateAt < 400) return;
      lastActivateAt = now;
      // 지도 롱프레스/클릭이 이어서 핀을 찍지 않도록
      ignoreClickUntil = now + 700;
      selectMarketByLabel(m.baseName);
    };
    // 라벨 위에서 터치가 시작되면 지도 롱프레스 취소
    content.addEventListener("touchstart", (e) => {
      // 두 손가락이면 라벨 처리 자체를 하지 않음
      if (e.touches.length >= 2) {
        multiTouchActive = true;
        lastMultiTouchAt = Date.now();
        return;
      }
      e.stopPropagation();
    }, { passive: true });
    content.addEventListener("mousedown", (e) => {
      e.stopPropagation();
    });
    content.addEventListener("click", onLabelActivate);
    content.addEventListener("touchend", (e) => {
      // 핀치 줌 중·직후 touchend 로 라벨이 눌리는 것 방지
      if (multiTouchActive || (e.touches && e.touches.length >= 1)) return;
      if (Date.now() - lastMultiTouchAt < 300) return;
      onLabelActivate(e);
    }, { passive: false });
    content.addEventListener("dblclick", (e) => {
      e.preventDefault();
      e.stopPropagation();
    });

    const overlay = new kakao.maps.CustomOverlay({
      map: labelVisible,
      position,
      content,
      yAnchor: 1,
      xAnchor: 0.5,
      zIndex: 21
    });

    marketLabelOverlays.push({ marketName: m.baseName, type: m.type, marker, overlay, content });
  });
}

/* =========================================================
   6. 검색 강조(빨간색) 레이어 - 구역 배경은 그대로 두고 그 위에 표시
   ========================================================= */
function clearHighlights() {
  highlightOverlays.forEach(o => {
    if (o.polygon) o.polygon.setMap(null);
    if (o.marker) o.marker.setMap(null);
  });
  highlightOverlays = [];
}

function drawHighlightPolygon(parcel) {
  const path = toLatLngPath(parcel.coords);
  const polygon = new kakao.maps.Polygon({
    map,
    path,
    strokeWeight: 3,
    strokeColor: COLOR_HIGHLIGHT,
    strokeOpacity: 1,
    fillColor: COLOR_HIGHLIGHT,
    fillOpacity: 0.55,
    zIndex: 10
  });
  highlightOverlays.push({ polygon });
  return path;
}

// 검색으로 찾은 특정 주소(들)를 빨간색으로 강조. 다른 구역들은 항상 표시된 배경(파란색)을 그대로 유지.
function focusOnParcels(parcels) {
  clearHighlights();

  // 라벨 선택/체크리스트 상태를 초기화해서 모든 구역이 다시 기본 색으로 보이게 함
  closeChecklist();
  selectedMarket = null;
  applyZoneColorState();

  const paths = parcels.map(p => drawHighlightPolygon(p));
  fitBoundsToPaths(paths);
}

/* =========================================================
   6-1. 우클릭(PC) / 꾹 누르기(모바일) - 위치(필지) 핀 찍기 (여러 개 가능)
   - 필지 안을 우클릭하면 해당 필지 중앙에 마커 생성
   - 이미 핀이 찍힌 필지 안을 클릭(또는 다시 우클릭)하면 그 핀 제거
   ========================================================= */

// 점-다각형 포함 여부 (ray casting)
function pointInPolygon(lat, lng, coords) {
  if (!coords || coords.length < 3) return false;
  let inside = false;
  for (let i = 0, j = coords.length - 1; i < coords.length; j = i++) {
    const yi = coords[i].lat, xi = coords[i].lng;
    const yj = coords[j].lat, xj = coords[j].lng;
    const intersect =
      ((yi > lat) !== (yj > lat)) &&
      (lng < ((xj - xi) * (lat - yi)) / ((yj - yi) || 1e-15) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
}

function findParcelAt(latlng) {
  const lat = latlng.getLat();
  const lng = latlng.getLng();
  let found = null;
  for (let i = 0; i < MAP_DATA.parcels.length; i++) {
    const p = MAP_DATA.parcels[i];
    if (pointInPolygon(lat, lng, p.coords)) found = p;
  }
  return found;
}

function findZoneAt(latlng) {
  const lat = latlng.getLat();
  const lng = latlng.getLng();
  let found = null;
  for (let i = 0; i < MAP_DATA.zones.length; i++) {
    const z = MAP_DATA.zones[i];
    if (pointInPolygon(lat, lng, z.coords)) found = z;
  }
  return found;
}

function getCentroidLatLng(coords) {
  let latSum = 0;
  let lngSum = 0;
  const n = coords.length || 1;
  coords.forEach((c) => {
    latSum += c.lat;
    lngSum += c.lng;
  });
  return new kakao.maps.LatLng(latSum / n, lngSum / n);
}

// 중심점 기준 원형 클릭 영역 좌표 (대략 radiusMeters)
function makeCircleCoords(centerLat, centerLng, radiusMeters, steps) {
  steps = steps || 28;
  const coords = [];
  const dLat = radiusMeters / 111320;
  const dLng = radiusMeters / (111320 * Math.cos((centerLat * Math.PI) / 180) || 1e-6);
  for (let i = 0; i <= steps; i++) {
    const a = (i / steps) * Math.PI * 2;
    coords.push({
      lat: centerLat + dLat * Math.sin(a),
      lng: centerLng + dLng * Math.cos(a)
    });
  }
  return coords;
}

function getAreaKey(key) {
  if (!key) return key;
  if (key.startsWith("free|")) return key;
  const parts = key.split("|");
  // parcel|market|address|lat|lng  /  zone|market|zoneNo|lat|lng
  if (parts[0] === "parcel" && parts.length >= 3) return `parcel|${parts[1]}|${parts[2]}`;
  if (parts[0] === "zone" && parts.length >= 3) return `zone|${parts[1]}|${parts[2]}`;
  return key;
}

function ensureAreaPolygon(areaKey, hitCoords) {
  if (pickedAreaPolygons[areaKey]) {
    pickedAreaPolygons[areaKey].refCount += 1;
    return pickedAreaPolygons[areaKey];
  }
  const polygon = new kakao.maps.Polygon({
    map,
    path: toLatLngPath(hitCoords),
    strokeWeight: 2,
    strokeColor: "#6b21a8",
    strokeOpacity: 0.95,
    fillColor: "#6b21a8",
    fillOpacity: 0.32,
    zIndex: 40
  });
  kakao.maps.event.addListener(polygon, "click", () => {
    removeLastPickedInArea(areaKey);
  });
  pickedAreaPolygons[areaKey] = { polygon, hitCoords, refCount: 1 };
  return pickedAreaPolygons[areaKey];
}

function releaseAreaPolygon(areaKey) {
  const entry = pickedAreaPolygons[areaKey];
  if (!entry) return;
  entry.refCount -= 1;
  if (entry.refCount <= 0) {
    if (entry.polygon) entry.polygon.setMap(null);
    delete pickedAreaPolygons[areaKey];
  }
}

function removeLastPickedInArea(areaKey) {
  for (let i = pickedLocations.length - 1; i >= 0; i--) {
    if (pickedLocations[i].areaKey === areaKey) {
      removePickedLocationById(pickedLocations[i].id);
      return;
    }
  }
}

function clearPickedLocation() {
  pickedLocations.forEach((item) => {
    if (item.marker) item.marker.setMap(null);
    if (item.overlay) item.overlay.setMap(null);
  });
  Object.keys(pickedAreaPolygons).forEach((k) => {
    if (pickedAreaPolygons[k].polygon) pickedAreaPolygons[k].polygon.setMap(null);
  });
  pickedAreaPolygons = {};
  pickedLocations = [];
  pickedViewIndex = -1;
}

function removePickedLocationById(id) {
  const idx = pickedLocations.findIndex((item) => item.id === id);
  if (idx < 0) return;

  const item = pickedLocations[idx];
  if (item.marker) item.marker.setMap(null);
  if (item.overlay) item.overlay.setMap(null);
  if (item.areaKey) releaseAreaPolygon(item.areaKey);
  pickedLocations.splice(idx, 1);

  if (pickedLocations.length === 0) {
    pickedViewIndex = -1;
    renderInitialOverview();
  } else {
    // 삭제 후 → 바로 이전 핀으로 이동 (없으면 다음=남은 목록의 첫 항목)
    // idx 삭제 전 기준: 이전 = idx-1, 없으면 0 (삭제 후 앞당겨진 다음 핀)
    let newIndex = idx - 1;
    if (newIndex < 0) newIndex = 0;
    if (newIndex >= pickedLocations.length) newIndex = pickedLocations.length - 1;
    showPickedLocationAt(newIndex, true); // panMap: 이전 위치로 지도 이동
  }
}

/** 우클릭 핀 목록에서 index번째 상세정보를 표시하고, panMap이면 지도 중심도 이동 */
function showPickedLocationAt(index, panMap) {
  if (!pickedLocations.length) return;
  const i = Math.max(0, Math.min(index, pickedLocations.length - 1));
  pickedViewIndex = i;
  const item = pickedLocations[i];
  renderPickedLocationDetails(
    item.latlng,
    item.jibunFull,
    item.roadFull,
    item.parcelAddress,
    pickedLocations.length,
    item.market,
    i
  );
  if (panMap && item.latlng && map) {
    map.panTo(item.latlng);
  }
}

function goPickedPrev() {
  if (pickedLocations.length <= 1) return;
  const next = (pickedViewIndex - 1 + pickedLocations.length) % pickedLocations.length;
  showPickedLocationAt(next, true);
}

function goPickedNext() {
  if (pickedLocations.length <= 1) return;
  const next = (pickedViewIndex + 1) % pickedLocations.length;
  showPickedLocationAt(next, true);
}

/** 구역명 클릭 시: 해당 상점가 구역으로 이동 + 강조 */
function focusMarketZone(marketName) {
  if (!marketName) return;
  clearHighlights();
  closeChecklist();
  selectedMarket = marketName;
  applyZoneColorState();
  openChecklist(marketName);
  const zones = getZonesForMarket(marketName);
  if (zones.length) {
    fitBoundsToPaths(zones.map(z => toLatLngPath(z.coords)));
  }
}

// 두 좌표 사이 대략 거리(m)
function distanceMeters(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// 클릭 지점에서 maxMeters 이내 가장 가까운 핀
function findNearestPicked(latlng, maxMeters) {
  const lat = latlng.getLat();
  const lng = latlng.getLng();
  let best = null;
  let bestD = maxMeters;
  for (let i = 0; i < pickedLocations.length; i++) {
    const item = pickedLocations[i];
    if (!item.latlng) continue;
    const d = distanceMeters(lat, lng, item.latlng.getLat(), item.latlng.getLng());
    if (d <= bestD) {
      bestD = d;
      best = item;
    }
  }
  return best;
}

// 저장된 hitCoords 기준으로, 클릭 지점이 포함된 핀을 찾아 제거 (왼쪽 클릭용)
function removePickedAtLatLng(latlng) {
  // 1) 마커 근처면 그 핀 제거
  const near = findNearestPicked(latlng, 25);
  if (near) {
    removePickedLocationById(near.id);
    return true;
  }
  // 2) 보라색 영역 안이면 그 영역의 가장 최근 핀 제거
  const lat = latlng.getLat();
  const lng = latlng.getLng();
  const areaKeys = Object.keys(pickedAreaPolygons);
  for (let i = areaKeys.length - 1; i >= 0; i--) {
    const entry = pickedAreaPolygons[areaKeys[i]];
    if (entry.hitCoords && pointInPolygon(lat, lng, entry.hitCoords)) {
      removeLastPickedInArea(areaKeys[i]);
      return true;
    }
  }
  return false;
}

// 전체 주소에서 맨 앞의 시/도 이름만 제거
function stripCityName(fullAddress, cityName) {
  if (!fullAddress) return "";
  if (cityName && fullAddress.startsWith(cityName)) {
    return fullAddress.slice(cityName.length).trim();
  }
  return fullAddress;
}

function handleMapPick(latlng) {
  // 우클릭한 좌표를 숫자로 고정 (항상 그 위치에 마커)
  const clickLat = latlng.getLat();
  const clickLng = latlng.getLng();
  const clickPos = new kakao.maps.LatLng(clickLat, clickLng);

  // 기존 마커를 거의 같은 자리(약 12m 이내)에서 우클릭하면 제거만 수행
  const near = findNearestPicked(clickPos, 12);
  if (near) {
    removePickedLocationById(near.id);
    return;
  }

  const parcel = findParcelAt(clickPos);
  const zone = findZoneAt(clickPos);

  // 1) 구역(또는 필지) 안 → 마커는 클릭 위치, 보라색은 구역 단위로 한 겹만
  if (zone || parcel) {
    const labelText = parcel
      ? (parcel.address || parcel.market || "선택한 위치")
      : (zone.market ? `${getMarketDisplayName(zone.market)} 구역${zone.zoneNo || ""}` : "선택한 구역");
    const jibunFull = parcel ? (parcel.address || "") : "";
    const parcelAddress = parcel ? (parcel.address || "") : (zone ? (zone.market || "") : "");
    // 보라색은 구역 좌표 우선 (필지마다 겹쳐 진해지지 않도록)
    const hitCoords = zone ? zone.coords : parcel.coords;
    const areaId = zone
      ? `zone|${zone.market}|${zone.zoneNo}`
      : `parcel|${parcel.market}|${parcel.address}`;

    addPickedPin({
      center: clickPos,
      hitCoords,
      labelText,
      jibunFull,
      roadFull: "",
      parcelAddress,
      market: zone ? zone.market : parcel.market,
      key: `${areaId}|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`,
      showPurple: true
    });
    return;
  }

  // 2) 구역 밖 → 마커만 (보라색 원 없음)
  if (!geocoder) {
    addPickedPin({
      center: clickPos,
      hitCoords: null,
      labelText: "선택한 위치",
      jibunFull: "",
      roadFull: "",
      parcelAddress: "",
      market: null,
      key: `free|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`,
      showPurple: false
    });
    return;
  }

  geocoder.coord2Address(clickLng, clickLat, (result, status) => {
    let jibunFull = "";
    let roadFull = "";
    let labelText = "선택한 위치";
    if (status === kakao.maps.services.Status.OK && result[0]) {
      if (result[0].address) {
        jibunFull = result[0].address.address_name;
        labelText = stripCityName(jibunFull, result[0].address.region_1depth_name) || jibunFull;
      }
      if (result[0].road_address) {
        roadFull = result[0].road_address.address_name;
      }
    }
    addPickedPin({
      center: clickPos,
      hitCoords: null,
      labelText,
      jibunFull,
      roadFull,
      parcelAddress: jibunFull,
      market: null,
      key: `free|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`,
      showPurple: false
    });
  });
}

// 지도/폴리곤 왼쪽 클릭 → 해당 영역/근처 핀 제거
function handleMapClick(latlng) {
  removePickedAtLatLng(latlng);
}

function addPickedPin({ center, hitCoords, labelText, jibunFull, roadFull, parcelAddress, market, key, showPurple, skipDetail }) {
  // 좌표를 다시 숫자로 복사해 새 LatLng 생성 (참조/변형 문제 방지)
  const pos = new kakao.maps.LatLng(center.getLat(), center.getLng());
  const id = ++pickedLocationIdSeq;
  const areaKey = getAreaKey(key);

  // 구역/필지일 때만 보라색 폴리곤 표시, 같은 구역은 폴리곤 1개만 유지
  if (showPurple && hitCoords && hitCoords.length) {
    ensureAreaPolygon(areaKey, hitCoords);
  }

  // 검색·우클릭 핀이 다른 라벨보다 앞에 보이도록 zIndex를 높게
  const marker = new kakao.maps.Marker({
    map,
    position: pos,
    zIndex: 120
  });

  const content = document.createElement("div");
  content.className = "market-label picked-label";
  content.textContent = labelText;
  content.title = showPurple
    ? "마커 또는 보라색 영역을 클릭하면 사라집니다"
    : "마커를 클릭하면 사라집니다";
  content.style.cursor = "pointer";

  const overlay = new kakao.maps.CustomOverlay({
    map,
    position: pos,
    content,
    yAnchor: 1,
    xAnchor: 0.5,
    zIndex: 121,
    clickable: true
  });

  const removeThis = () => removePickedLocationById(id);

  kakao.maps.event.addListener(marker, "click", removeThis);
  content.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    removeThis();
  });

  pickedLocations.push({
    id,
    marker,
    overlay,
    areaKey: (showPurple && hitCoords) ? areaKey : null,
    hitCoords: hitCoords || null,
    latlng: pos,
    jibunFull: jibunFull || "",
    roadFull: roadFull || "",
    parcelAddress: parcelAddress || "",
    market: market || null,
    parcelKey: key
  });

  // 검색 결과용 핀은 상세 패널을 덮어쓰지 않음
  if (!skipDetail) {
    pickedViewIndex = pickedLocations.length - 1;
    showPickedLocationAt(pickedViewIndex, false);
  }
}

/** 검색 결과 한 건을 우클릭과 동일하게 핀+보라 라벨로 표시 */
function pinFromSearchItem(item, options = {}) {
  if (!item || !item.latlng) return;
  const parcel = item.parcelId != null
    ? MAP_DATA.parcels.find((p) => p.id === item.parcelId)
    : findParcelAt(item.latlng);
  const zone = findZoneAt(item.latlng);
  const hitCoords = (parcel && parcel.coords) || (zone && zone.coords) || null;
  const inZone = !!(parcel || zone);
  const labelText = item.placeName || item.parcelAddress || item.jibunFull || "검색 위치";

  addPickedPin({
    center: item.latlng,
    hitCoords,
    labelText,
    jibunFull: item.jibunFull || "",
    roadFull: item.roadFull || "",
    parcelAddress: item.parcelAddress || "",
    market: item.market || (parcel && parcel.market) || (zone && zone.market) || null,
    key: item.parcelId != null
      ? `search|parcel|${item.parcelId}`
      : `search|${item.latlng.getLat().toFixed(6)}|${item.latlng.getLng().toFixed(6)}`,
    showPurple: inZone && !!hitCoords,
    skipDetail: options.skipDetail !== false
  });
}

/** 검색 위치로 지도 확대 (상호·주소 공통) */
function zoomToSearchPoint(latlng) {
  if (!map || !latlng) return;
  map.setCenter(latlng);
  map.setLevel(2);
}

/**
 * 검색·우클릭 공통 상세 패널
 * titleBase: "선택한 위치 상세정보" | "검색 결과"
 * items: [{ latlng, jibunFull, roadFull, market }]
 * index: 현재 인덱스
 * onPrev / onNext: 네비게이션 콜백
 */
function renderUnifiedDetailPanel({ titleBase, items, index, onPrev, onNext }) {
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");
  const list = document.getElementById("resultList");

  const n = items.length;
  const idx = Math.max(0, Math.min(index, n - 1));
  const item = items[idx] || {};
  const displayNum = n > 0 ? (idx + 1) : 0;

  badge.style.display = "none";

  if (n > 1) {
    title.innerHTML = `
      <span class="picked-title-text">${titleBase} (${displayNum}/${n})</span>
      <span class="picked-nav" role="group" aria-label="위치 이동">
        <button type="button" class="picked-nav-btn" id="unifiedNavPrev" title="이전">◀</button>
        <button type="button" class="picked-nav-btn" id="unifiedNavNext" title="다음">▶</button>
      </span>
    `;
    const prevBtn = document.getElementById("unifiedNavPrev");
    const nextBtn = document.getElementById("unifiedNavNext");
    if (prevBtn && onPrev) prevBtn.onclick = (e) => { e.stopPropagation(); onPrev(); };
    if (nextBtn && onNext) nextBtn.onclick = (e) => { e.stopPropagation(); onNext(); };
  } else {
    title.textContent = titleBase;
  }

  const zoneNotice = buildZoneNoticeHtml(item.market ? [item.market] : []);
  list.innerHTML = zoneNotice + buildDetailRowsHtml(item.latlng, item.jibunFull, item.roadFull, item.placeName);
  bindZoneLinkClicks(list);
}

function renderPickedLocationDetails(latlng, jibunFull, roadFull, parcelAddress, count, market, viewIndex) {
  // 검색 상세 상태는 우클릭 상세가 열리면 비움
  searchDetailItems = [];
  searchDetailIndex = 0;

  const n = count != null ? count : pickedLocations.length;
  const idx = viewIndex != null ? viewIndex : (pickedViewIndex >= 0 ? pickedViewIndex : Math.max(0, n - 1));

  const items = pickedLocations.map(p => ({
    latlng: p.latlng,
    jibunFull: p.jibunFull,
    roadFull: p.roadFull,
    market: p.market
  }));

  // 단건일 때는 인자로 받은 값을 우선 사용 (직후 push된 항목)
  if (items.length === 0 && latlng) {
    items.push({ latlng, jibunFull, roadFull, market });
  }

  renderUnifiedDetailPanel({
    titleBase: "선택한 위치 상세정보",
    items: items.length ? items : [{ latlng, jibunFull, roadFull, market }],
    index: idx,
    onPrev: goPickedPrev,
    onNext: goPickedNext
  });
}

/* =========================================================
   7. 결과 패널 렌더링
   ========================================================= */
function renderResultList(parcels) {
  const list = document.getElementById("resultList");
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");

  if (!parcels.length) {
    searchDetailItems = [];
    searchDetailIndex = 0;
    badge.style.display = "";
    list.innerHTML = `<div class="result-empty">일치하는 결과가 없습니다. 시장명, 지번 주소, 또는 상호명(예: 진성 아구찜)으로 검색해보세요.</div>`;
    title.textContent = "검색 결과";
    badge.textContent = "골목형상점가 0곳";
    return;
  }

  // 빨간 강조 대신 우클릭과 동일하게 핀+보라 라벨
  clearHighlights();
  clearPickedLocation();

  searchDetailItems = parcels.map(p => {
    const center = getParcelCenter(p);
    return {
      latlng: center,
      jibunFull: p.address || "",
      roadFull: p.roadAddress || "",
      parcelAddress: formatParcelAddress(p),
      market: p.market || null,
      parcelId: p.id,
      placeName: null
    };
  });
  searchDetailItems.forEach((item) => pinFromSearchItem(item, { skipDetail: true }));
  searchDetailIndex = 0;
  if (searchDetailItems[0] && searchDetailItems[0].latlng) {
    zoomToSearchPoint(searchDetailItems[0].latlng);
  }
  showSearchDetailAt(0, false);
}

function showSearchDetailAt(index, panMap) {
  if (!searchDetailItems.length) return;
  const i = Math.max(0, Math.min(index, searchDetailItems.length - 1));
  searchDetailIndex = i;
  const item = searchDetailItems[i];

  renderUnifiedDetailPanel({
    titleBase: "검색 결과",
    items: searchDetailItems,
    index: i,
    onPrev: goSearchPrev,
    onNext: goSearchNext
  });

  if (panMap && item.latlng && map) {
    zoomToSearchPoint(item.latlng);
  }
}

function goSearchPrev() {
  if (searchDetailItems.length <= 1) return;
  const next = (searchDetailIndex - 1 + searchDetailItems.length) % searchDetailItems.length;
  showSearchDetailAt(next, true);
}

function goSearchNext() {
  if (searchDetailItems.length <= 1) return;
  const next = (searchDetailIndex + 1) % searchDetailItems.length;
  showSearchDetailAt(next, true);
}

// 초기 로딩 시: 결과 패널에 전체 골목형상점가 목록을 보여줌 (지도에는 이미 구역 배경이 항상 표시되어 있음)
// 유형 표시명 (제목/배지용)
const TYPE_LABEL_ORDER = ["전통시장", "상점가", "골목형상점가"];
const TYPE_LABEL_SHORT = {
  "전통시장": "전통시장",
  "상점가": "상점가",
  "골목형상점가": "골목형 상점가"
};

// 현재 켜져 있는 유형 기준으로 결과 패널(제목·배지·목록)을 갱신
function renderInitialOverview() {
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");
  const list = document.getElementById("resultList");
  badge.style.display = "";

  const activeOrdered = TYPE_LABEL_ORDER.filter(t => activeTypes.has(t));
  const visibleMarkets = MAP_DATA.markets.filter(m => activeTypes.has(m.type));

  if (!activeOrdered.length) {
    title.textContent = "선택된 유형 없음";
    badge.textContent = "0곳";
    list.innerHTML = `<div class="result-empty">지도 표시 기준에서 유형을 하나 이상 선택해주세요.</div>`;
    return;
  }

  const titleNames = activeOrdered.map(t => TYPE_LABEL_SHORT[t] || t).join(", ");
  title.textContent = `${titleNames} 전체 구역도`;

  if (activeOrdered.length === 1) {
    const t = activeOrdered[0];
    const cnt = visibleMarkets.filter(m => m.type === t).length;
    badge.textContent = `${TYPE_LABEL_SHORT[t] || t} ${cnt}곳`;
  } else {
    badge.textContent = `총 ${visibleMarkets.length}곳`;
  }

  if (!visibleMarkets.length) {
    list.innerHTML = `<div class="result-empty">표시할 데이터가 없습니다.</div>`;
    return;
  }

  // 유형 순서대로 정렬 후 목록 표시
  const typeRank = Object.fromEntries(TYPE_LABEL_ORDER.map((t, i) => [t, i]));
  const sorted = [...visibleMarkets].sort((a, b) => {
    const ra = typeRank[a.type] ?? 99;
    const rb = typeRank[b.type] ?? 99;
    if (ra !== rb) return ra - rb;
    return (a.name || "").localeCompare(b.name || "", "ko");
  });

  list.innerHTML = sorted.map(m => {
    const zoneCount = getZonesForMarket(m.baseName).length;
    const typeLabel = TYPE_LABEL_SHORT[m.type] || m.type;
    return `
      <div class="result-item" data-market="${m.baseName}" data-type="${m.type}">
        <span class="r-market">${getMarketLabelText(m)}</span>
        <span class="r-addr">${typeLabel} · 구역 ${zoneCount}개</span>
      </div>
    `;
  }).join("");

  list.querySelectorAll(".result-item").forEach(el => {
    el.addEventListener("click", () => {
      const marketName = el.dataset.market;
      // 체크리스트 + 구역 강조
      clearHighlights();
      closeChecklist();
      selectedMarket = marketName;
      applyZoneColorState();
      openChecklist(marketName);
      // 해당 상점가 구역으로 지도 이동
      const zones = getZonesForMarket(marketName);
      if (zones.length) {
        fitBoundsToPaths(zones.map(z => toLatLngPath(z.coords)));
      }
    });
  });
}

// 검색 중이 아닐 때만 유형 필터 변경에 맞춰 결과 패널을 개요 모드로 갱신
function maybeRefreshOverviewPanel() {
  const q = (document.getElementById("searchInput")?.value || "").trim();
  if (q) return; // 검색어가 있으면 검색 결과 유지
  if (highlightOverlays.length) return; // 검색 강조 중이면 유지
  renderInitialOverview();
}

/* =========================================================
   8. 검색 처리
   ========================================================= */

/** 카카오 장소 결과가 세종특별자치시인지 엄격 판별 (대전·공주·충남 등 제외) */
function isSejongPlace(place) {
  if (!place) return false;
  const addr = `${place.address_name || ""} ${place.road_address_name || ""}`;

  // 타 시·군 주소는 명시적으로 제외 (세종 문구가 상호에만 있어도 탈락)
  if (/대전|공주|천안|청주|아산|논산|계룡|금산|부여|서산|당진|보령|홍성|예산|청양|서천|태안|충청남도|충남\s/.test(addr)) {
    // 주소에 「세종특별자치시」가 함께 있는 경우만 예외 (실제로는 거의 없음)
    if (!/세종특별자치시/.test(addr)) return false;
  }

  // 지번·도로명 주소에 세종시/세종특별자치시가 있어야 함
  if (/세종특별자치시|세종시/.test(addr)) return true;
  // "세종 보람동 …" 형태 (카카오 축약 표기)
  if (/(^|\s)세종(\s|시)/.test(addr)) return true;

  return false;
}

/**
 * 세종시 한정 상호(장소) 키워드 검색
 * - 사용자는 「진성 아구찜」만 입력해도 됨
 * - 여러 방식으로 재시도 후 세종 결과만 반환
 * - callback(places, meta) meta: { status, tried }
 */
function searchPlacesInSejong(query, callback) {
  if (!placesService) {
    callback([], { status: "NO_SERVICE", tried: [] });
    return;
  }
  const q = (query || "").trim();
  if (!q) {
    callback([], { status: "EMPTY_QUERY", tried: [] });
    return;
  }

  const center = new kakao.maps.LatLng(SEJONG_CENTER_LAT, SEJONG_CENTER_LNG);
  const options = {
    location: center,
    radius: SEJONG_SEARCH_RADIUS_M,
    size: 15
  };

  // 공백 유무·세종 접두를 바꿔가며 재시도 (항상 세종 중심 반경 안에서만)
  const keywordVariants = [];
  const noSpace = q.replace(/\s+/g, "");
  keywordVariants.push(q);
  if (noSpace !== q) keywordVariants.push(noSpace);
  if (!/세종/.test(q)) {
    keywordVariants.push(`세종 ${q}`);
    if (noSpace !== q) keywordVariants.push(`세종 ${noSpace}`);
    keywordVariants.push(`세종특별자치시 ${q}`);
  }

  const tried = [];
  let lastStatus = null;

  function runNext(i) {
    if (i >= keywordVariants.length) {
      callback([], { status: lastStatus || "ZERO_RESULT", tried });
      return;
    }
    const keyword = keywordVariants[i];
    tried.push(keyword);

    placesService.keywordSearch(
      keyword,
      (data, status) => {
        lastStatus = status;
        if (status === kakao.maps.services.Status.OK && data && data.length) {
          const filtered = data.filter(isSejongPlace);
          if (filtered.length) {
            callback(filtered, { status, tried });
            return;
          }
        }
        runNext(i + 1);
      },
      options
    );
  }

  runNext(0);
}

/** 장소 검색 결과 → 우클릭과 동일하게 핀+보라 라벨 + 상세 패널 */
function renderPlaceSearchResults(places, originalQuery) {
  if (!places.length) return;

  clearHighlights();
  closeChecklist();
  selectedMarket = null;
  applyZoneColorState();
  clearPickedLocation();

  const detailItems = places.map((place) => {
    const lat = parseFloat(place.y);
    const lng = parseFloat(place.x);
    const coords = new kakao.maps.LatLng(lat, lng);
    const parcel = findParcelAt(coords);
    const zone = findZoneAt(coords);
    const market = (parcel && parcel.market) || (zone && zone.market) || null;
    return {
      latlng: coords,
      jibunFull: place.address_name || "",
      roadFull: place.road_address_name || "",
      parcelAddress: place.address_name || place.place_name || originalQuery,
      market,
      placeName: place.place_name || originalQuery,
      parcelId: parcel ? parcel.id : null
    };
  });

  // 모든 결과에 핀+보라 라벨 (우클릭과 동일)
  detailItems.forEach((item) => pinFromSearchItem(item, { skipDetail: true }));

  searchDetailItems = detailItems;
  searchDetailIndex = 0;
  if (detailItems[0] && detailItems[0].latlng) {
    zoomToSearchPoint(detailItems[0].latlng);
  }
  showSearchDetailAt(0, false);
}

function showSearchNoResult(message) {
  searchDetailItems = [];
  searchDetailIndex = 0;
  const list = document.getElementById("resultList");
  list.innerHTML = `<div class="result-empty">${message || "검색 결과가 없습니다. 시장명, 주소, 또는 상호명을 입력해주세요."}</div>`;
  document.getElementById("resultTitle").textContent = "검색 결과";
  const badge = document.getElementById("alleyCountBadge");
  badge.style.display = "";
  badge.textContent = "골목형상점가 0곳";
}

function geocodeFallbackSearch(query) {
  if (!geocoder) {
    showSearchNoResult();
    return;
  }
  geocoder.addressSearch(query, (result, status) => {
    if (status === kakao.maps.services.Status.OK && result[0]) {
      const coords = new kakao.maps.LatLng(result[0].y, result[0].x);
      const clickLat = coords.getLat();
      const clickLng = coords.getLng();

      clearHighlights();
      closeChecklist();
      selectedMarket = null;
      applyZoneColorState();

      // 검색된 좌표가 등록된 구역/필지 안에 있는지 한 번 더 확인
      const parcel = findParcelAt(coords);
      const zone = findZoneAt(coords);

      if (zone || parcel) {
        // 데이터 시트에는 주소 문자열이 없었지만 좌표상 구역 안인 경우
        const matchedParcels = parcel ? [parcel] : MAP_DATA.parcels.filter(p => p.market === zone.market);
        if (matchedParcels.length) {
          renderResultList(matchedParcels);
          return;
        }
      }

      let jibunFull = "";
      let roadFull = "";
      let labelText = query;
      if (result[0].address) {
        jibunFull = result[0].address.address_name || "";
        labelText = stripCityName(jibunFull, result[0].address.region_1depth_name) || jibunFull || query;
      }
      if (result[0].road_address) {
        roadFull = result[0].road_address.address_name || "";
      }

      const hitCoords = (parcel && parcel.coords) || (zone && zone.coords) || null;
      const market = (parcel && parcel.market) || (zone && zone.market) || null;

      // 우클릭/꾹 누르기와 동일: 핀 + 보라 라벨, 확대
      clearPickedLocation();
      addPickedPin({
        center: coords,
        hitCoords,
        labelText,
        jibunFull,
        roadFull,
        parcelAddress: jibunFull || query,
        market,
        key: `free|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`,
        showPurple: !!(hitCoords && market)
      });
      zoomToSearchPoint(coords);
    } else {
      showSearchNoResult("검색 결과가 없습니다. 시장명, 주소, 또는 상호명을 입력해주세요.");
    }
  });
}

function dismissSearchKeyboard() {
  const input = document.getElementById("searchInput");
  if (input) input.blur();
}

function handleSearch() {
  const query = document.getElementById("searchInput").value;
  if (!query.trim()) return;

  // 모바일: 검색 실행 시 키보드(이동/검색 버튼) 닫기
  dismissSearchKeyboard();

  const q = query.trim();
  // 주소처럼 보이면 필지·주소 매칭 우선, 아니면 시장명·상호명 검색
  const looksLikeAddress = /\d/.test(q) || q.includes("-") || q.includes("로") || q.includes("길");
  const matched = looksLikeAddress ? searchParcelsExact(q) : searchParcels(q);

  if (matched.length) {
    // 빨간 강조 없이 핀+보라 라벨로 통일 (renderResultList 내부에서 처리)
    renderResultList(matched);
    return;
  }

  if (!map) return;

  // 1) 상호명(장소) 검색 — 세종시 한정 (「세종」은 자동 처리)
  // 2) 없으면 주소 지오코딩 (상호명은 보통 주소 검색에 안 걸리므로 안내 문구 구분)
  searchPlacesInSejong(q, (places, meta) => {
    if (places && places.length) {
      renderPlaceSearchResults(places, q);
      return;
    }

    // 주소처럼 보이면 지오코딩 시도
    if (looksLikeAddress) {
      const geoQuery = /세종/.test(q) ? q : `세종 ${q}`;
      geocodeFallbackSearch(geoQuery);
      return;
    }

    // 상호 검색 실패 안내 (API 오류 vs 결과 없음)
    if (meta && meta.status === kakao.maps.services.Status.ERROR) {
      showSearchNoResult(
        "상호 검색에 실패했습니다. 카카오 개발자 콘솔에서 카카오맵(로컬) API가 활성화되어 있는지, 도메인이 등록되어 있는지 확인해주세요."
      );
      return;
    }
    if (meta && meta.status === "NO_SERVICE") {
      showSearchNoResult("장소 검색 서비스를 불러오지 못했습니다. 페이지를 새로고침한 뒤 다시 시도해주세요.");
      return;
    }

    showSearchNoResult(
      `「${q}」에 해당하는 세종시 상호를 찾지 못했습니다. 카카오맵에 등록된 이름과 같은지 확인하거나, 지번·도로명 주소로 검색해보세요.`
    );
  });
}

document.getElementById("searchBtn").addEventListener("click", handleSearch);
document.getElementById("searchInput").addEventListener("keydown", (e) => {
  if (e.key === "Enter") handleSearch();
});

/* -------- 결과 목록 접기/펼치기 (화살표만) -------- */
(function setupPanelCollapse() {
  const sidePanel = document.getElementById("sidePanel");
  const btnBottom = document.getElementById("panelToggleBottom");
  const btnSide = document.getElementById("panelToggleSide");
  if (!sidePanel) return;

  function setCollapsed(collapsed) {
    sidePanel.classList.toggle("is-collapsed", collapsed);
    [btnBottom, btnSide].forEach((btn) => {
      if (!btn) return;
      btn.setAttribute("aria-expanded", collapsed ? "false" : "true");
      btn.setAttribute("aria-label", collapsed ? "목록 펼치기" : "목록 접기");
      btn.title = collapsed ? "목록 펼치기" : "목록 접기";
      const icon = btn.querySelector(".panel-toggle-icon");
      // 모바일: ▲ 접기 / ▼ 펼치기  |  PC: ◀ 접기 / ▶ 펼치기
      if (icon) {
        if (btn === btnBottom) icon.textContent = collapsed ? "▼" : "▲";
        else icon.textContent = collapsed ? "▶" : "◀";
      }
    });
    if (map && typeof map.relayout === "function") {
      setTimeout(() => {
        try { map.relayout(); } catch (_) {}
      }, 50);
    }
  }

  const toggle = () => setCollapsed(!sidePanel.classList.contains("is-collapsed"));
  if (btnBottom) btnBottom.addEventListener("click", toggle);
  if (btnSide) btnSide.addEventListener("click", toggle);
})();

/* -------- 지도 조이스틱 (터치·마우스·키보드 화살표) -------- */
(function setupMapJoystick() {
  const root = document.getElementById("mapJoystick");
  const knob = document.getElementById("mapJoystickKnob");
  if (!root || !knob) return;

  // 노브 최대 이동량: 조이스틱 크기에 비례
  const getMaxKnob = () => Math.max(12, root.offsetWidth / 2 - 16);
  // 기본 속도 / 컴퓨터(960px+) 는 1.5배
  const basePanSpeed = 4.5;
  const getPanSpeed = () => {
    const desktop = window.matchMedia && window.matchMedia("(min-width: 960px)").matches;
    const mobile = window.matchMedia && window.matchMedia("(max-width: 599px)").matches;
    if (desktop) return basePanSpeed * 2.7; // 기존 PC 속도의 1.8배
    if (mobile) return basePanSpeed * 1.5; // 기존 핸드폰 속도의 1.5배
    return basePanSpeed; // 패드 유지
  };

  let vecX = 0; // -1 ~ 1
  let vecY = 0;
  let keyVecX = 0;
  let keyVecY = 0;
  let pointerId = null;
  let rafId = null;

  function setKnob(nx, ny) {
    // nx, ny: -1 ~ 1
    const len = Math.hypot(nx, ny);
    if (len > 1) {
      nx /= len;
      ny /= len;
    }
    vecX = nx;
    vecY = ny;
    const maxK = getMaxKnob();
    knob.style.transform = `translate(${nx * maxK}px, ${ny * maxK}px)`;
    root.classList.toggle("is-active", Math.hypot(nx, ny) > 0.05);
  }

  function resetKnob() {
    setKnob(0, 0);
  }

  function applyCombinedVector() {
    // 조이스틱 + 키보드 합산 후 정규화
    let x = vecX + keyVecX;
    let y = vecY + keyVecY;
    const len = Math.hypot(x, y);
    if (len > 1) {
      x /= len;
      y /= len;
    }
    return { x, y };
  }

  function tick() {
    rafId = null;
    if (!map) {
      schedule();
      return;
    }
    const { x, y } = applyCombinedVector();
    if (Math.hypot(x, y) < 0.04) {
      schedule();
      return;
    }
    try {
      const level = map.getLevel();
      // 줌 아웃일수록 더 많이 이동 / PC는 1.5배
      const step = getPanSpeed() * (0.6 + level * 0.35);
      // 화면 기준: x>0 오른쪽, y>0 아래 → 지도 중심은 반대
      const proj = map.getProjection();
      const center = map.getCenter();
      const pt = proj.containerPointFromCoords(center);
      const next = new kakao.maps.Point(pt.x + x * step, pt.y + y * step);
      map.setCenter(proj.coordsFromContainerPoint(next));
    } catch (_) {}
    schedule();
  }

  function schedule() {
    if (rafId == null) rafId = requestAnimationFrame(tick);
  }
  schedule();

  function pointFromEvent(e) {
    const rect = root.getBoundingClientRect();
    const cx = rect.left + rect.width / 2;
    const cy = rect.top + rect.height / 2;
    const clientX = e.clientX != null ? e.clientX : (e.touches && e.touches[0] ? e.touches[0].clientX : cx);
    const clientY = e.clientY != null ? e.clientY : (e.touches && e.touches[0] ? e.touches[0].clientY : cy);
    const dx = clientX - cx;
    const dy = clientY - cy;
    const max = rect.width / 2 - 4;
    return { nx: dx / max, ny: dy / max };
  }

  function onPointerDown(e) {
    if (e.button != null && e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    pointerId = e.pointerId != null ? e.pointerId : "touch";
    try { root.setPointerCapture(e.pointerId); } catch (_) {}
    const p = pointFromEvent(e);
    setKnob(p.nx, p.ny);
  }

  function onPointerMove(e) {
    if (pointerId == null) return;
    if (e.pointerId != null && e.pointerId !== pointerId) return;
    e.preventDefault();
    const p = pointFromEvent(e);
    setKnob(p.nx, p.ny);
  }

  function onPointerUp(e) {
    if (pointerId == null) return;
    if (e.pointerId != null && e.pointerId !== pointerId) return;
    pointerId = null;
    resetKnob();
  }

  root.addEventListener("pointerdown", onPointerDown);
  root.addEventListener("pointermove", onPointerMove);
  root.addEventListener("pointerup", onPointerUp);
  root.addEventListener("pointercancel", onPointerUp);
  // 터치 폴백
  root.addEventListener("touchstart", (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!e.touches.length) return;
    pointerId = "touch";
    const p = pointFromEvent(e.touches[0]);
    setKnob(p.nx, p.ny);
  }, { passive: false });
  root.addEventListener("touchmove", (e) => {
    if (pointerId == null) return;
    e.preventDefault();
    if (!e.touches.length) return;
    const p = pointFromEvent(e.touches[0]);
    setKnob(p.nx, p.ny);
  }, { passive: false });
  root.addEventListener("touchend", () => {
    pointerId = null;
    resetKnob();
  });
  root.addEventListener("touchcancel", () => {
    pointerId = null;
    resetKnob();
  });

  // 키보드 화살표 연동
  const keyMap = {
    ArrowUp: { x: 0, y: -1 },
    ArrowDown: { x: 0, y: 1 },
    ArrowLeft: { x: -1, y: 0 },
    ArrowRight: { x: 1, y: 0 }
  };

  window.addEventListener("keydown", (e) => {
    if (!keyMap[e.key]) return;
    // 검색창 입력 중이면 무시
    const tag = (e.target && e.target.tagName) || "";
    if (tag === "INPUT" || tag === "TEXTAREA" || (e.target && e.target.isContentEditable)) return;
    e.preventDefault();
    const v = keyMap[e.key];
    if (v.x < 0) keyVecX = -1;
    if (v.x > 0) keyVecX = 1;
    if (v.y < 0) keyVecY = -1;
    if (v.y > 0) keyVecY = 1;
    // 키보드만 쓸 때도 노브 표시
    if (pointerId == null) setKnob(keyVecX, keyVecY);
  });

  window.addEventListener("keyup", (e) => {
    if (!keyMap[e.key]) return;
    const v = keyMap[e.key];
    if (v.x < 0 && keyVecX < 0) keyVecX = 0;
    if (v.x > 0 && keyVecX > 0) keyVecX = 0;
    if (v.y < 0 && keyVecY < 0) keyVecY = 0;
    if (v.y > 0 && keyVecY > 0) keyVecY = 0;
    if (pointerId == null) setKnob(keyVecX, keyVecY);
  });
})();

/* -------- 초기화 버튼: 사이트 첫 진입 상태로 복귀 -------- */
function resetToInitialView() {
  document.getElementById("searchInput").value = "";
  clearHighlights();
  clearPickedLocation();
  searchDetailItems = [];
  searchDetailIndex = 0;
  closeChecklist();
  selectedMarket = null;

  activeTypes = new Set(Object.keys(TYPE_COLORS));
  document.querySelectorAll(".type-chip").forEach(chip => chip.classList.add("active"));
  applyTypeVisibility();

  applyZoneColorState();
  renderInitialOverview();

  if (map) {
    map.setCenter(new kakao.maps.LatLng(36.479934, 127.286740));
    map.setLevel(4);
  }
}

document.getElementById("resetBtn").addEventListener("click", resetToInitialView);

/* -------- 지도를 볼 기준: 전통시장/상점가/골목형상점가 유형별 필터 -------- */
// activeTypes에 들어있는 유형만 구역/라벨/마커를 지도에 표시
// 라벨은 추가로 줌 레벨 조건도 만족해야 표시
function applyTypeVisibility() {
  Object.entries(zoneOverlaysByMarket).forEach(([marketName, polygons]) => {
    const type = getMarketType(marketName);
    const visible = activeTypes.has(type) ? map : null;
    polygons.forEach(p => p.setMap(visible));
  });

  updateLabelVisibilityByZoom();
}

// 유형별 개수를 칩 버튼 텍스트에 반영 (예: 골목형 상점가(12))
function updateTypeChipCounts() {
  const counts = { "전통시장": 0, "상점가": 0, "골목형상점가": 0 };
  MAP_DATA.markets.forEach(m => {
    if (counts[m.type] !== undefined) counts[m.type] += 1;
  });

  const labelMap = {
    "전통시장": "전통시장",
    "상점가": "상점가",
    "골목형상점가": "골목형 상점가"
  };

  document.querySelectorAll(".type-chip").forEach(chip => {
    const type = chip.dataset.type;
    const count = counts[type] || 0;
    const baseLabel = labelMap[type] || type;
    chip.textContent = `${baseLabel}(${count})`;
  });
}

document.querySelectorAll(".type-chip").forEach(chip => {
  chip.addEventListener("click", () => {
    const type = chip.dataset.type;
    const nowActive = !chip.classList.contains("active");
    chip.classList.toggle("active", nowActive);

    if (nowActive) {
      activeTypes.add(type);
    } else {
      activeTypes.delete(type);
      // 꺼진 유형이 현재 선택/체크리스트 대상이면 함께 정리
      if (selectedMarket && getMarketType(selectedMarket) === type) {
        closeChecklist();
        selectedMarket = null;
        applyZoneColorState();
      }
    }

    applyTypeVisibility();
    maybeRefreshOverviewPanel();
  });
});

/* -------- 마이크 음성 검색 -------- */
const SpeechRecognitionCtor = window.SpeechRecognition || window.webkitSpeechRecognition;
let recognition = null;
let isListening = false;

if (SpeechRecognitionCtor) {
  recognition = new SpeechRecognitionCtor();
  recognition.lang = "ko-KR";
  recognition.continuous = false;
  recognition.interimResults = false;

  recognition.onstart = () => {
    isListening = true;
    document.getElementById("micBtn").classList.add("listening");
  };

  recognition.onend = () => {
    isListening = false;
    document.getElementById("micBtn").classList.remove("listening");
  };

  recognition.onerror = () => {
    isListening = false;
    document.getElementById("micBtn").classList.remove("listening");
  };

  recognition.onresult = (event) => {
    const transcript = event.results[0][0].transcript.trim();
    if (transcript) {
      document.getElementById("searchInput").value = transcript;
      handleSearch();
    }
  };
} else {
  const micBtn = document.getElementById("micBtn");
  micBtn.disabled = true;
  micBtn.title = "이 브라우저는 음성 검색을 지원하지 않습니다.";
  micBtn.style.opacity = "0.4";
  micBtn.style.cursor = "not-allowed";
}

document.getElementById("micBtn").addEventListener("click", () => {
  if (!recognition) return;
  if (isListening) {
    recognition.stop();
  } else {
    try {
      recognition.start();
    } catch (e) {
      // 이미 시작된 상태 등 예외는 무시
    }
  }
});

/* =========================================================
   9. 초기화
   ========================================================= */

// 검색 결과 패널에 에러 메시지를 표시 (엑셀/카카오 로딩 실패 시, '불러오는 중' 문구가
// 그대로 남아있지 않도록 반드시 이 함수로 갱신함)
function showLoadError(message) {
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");
  const list = document.getElementById("resultList");
  if (title) title.textContent = "불러오기 실패";
  if (badge) badge.style.display = "none";
  if (list) list.innerHTML = `<div class="result-empty">⚠️ ${message}</div>`;
}

Promise.allSettled([loadKakaoSdk(KAKAO_APP_KEY), loadExcelData(EXCEL_FILE_URL)])
  .then(([kakaoResult, dataResult]) => {
    // 1) 카카오 SDK 자체가 실패하면 지도 영역에도 안내하고 여기서 종료
    if (kakaoResult.status === "rejected") {
      const err = kakaoResult.reason;
      const msg = (err && err.message === "NO_KEY")
        ? "카카오 지도 API 키가 설정되지 않았습니다."
        : "카카오 지도 SDK를 불러오지 못했습니다. (네트워크 상태 또는 도메인 등록을 확인해주세요)";
      showFallback(msg);
      showLoadError(msg);
      return;
    }

    // 2) 카카오 지도는 켰지만, 엑셀 데이터를 못 읽은 경우 - 지도는 표시하되 에러를 명확히 안내
    if (dataResult.status === "rejected") {
      const err = dataResult.reason;
      let msg;
      if (err && err.message === "EXCEL_FETCH_FAIL") {
        msg = `엑셀 데이터 파일(${EXCEL_FILE_URL})을 찾을 수 없습니다. index.html과 같은 폴더에 파일이 있는지, ` +
              `파일명 대소문자까지 정확한지 확인해주세요. file://로 직접 열었다면 로컬 서버(예: ` +
              `python3 -m http.server)나 GitHub Pages 같은 웹 서버로 열어야 합니다.`;
      } else {
        msg = `엑셀 데이터를 읽는 중 오류가 발생했습니다: ${err && err.message ? err.message : err}`;
      }
      initMap();
      showLoadError(msg);
      return;
    }

    // 3) 정상 로드
    MAP_DATA = dataResult.value;
    initMap();
    drawAllZonesBase();   // 유형별 색으로 모든 구역 표시
    drawMarketLabels();   // 각 구역 최고 위도 지점에 명칭 라벨 표시
    updateTypeChipCounts(); // 유형별 개수 표시
    renderInitialOverview();
    // 초기 화면: 지정 중심 + 줌 레벨 4 (fitBounds 사용하지 않음)
    map.setCenter(new kakao.maps.LatLng(36.479934, 127.286740));
    map.setLevel(4);
  });
