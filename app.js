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
let map, geocoder;
let MAP_DATA = { markets: [], parcels: [], zones: [] };

let zoneOverlaysByMarket = {};  // marketName -> [kakao.maps.Polygon, ...] (항상 유지되는 구역 배경)
let marketLabelOverlays = [];   // { marketName, marker, overlay, content } (항상 유지되는 라벨)
let highlightOverlays = [];     // { polygon, marker } (검색 시에만 갱신/삭제되는 강조 표시)
let pickedLocations = []; // [{ id, marker, overlay, latlng, jibunFull, roadFull, areaKey }] (우클릭/꾹 누르기 핀, 여러 개 가능)
let pickedLocationIdSeq = 0;
let pickedAreaPolygons = {}; // areaKey -> { polygon, hitCoords, refCount } (같은 구역 보라색은 하나만)
let selectedMarket = null;      // 라벨 클릭으로 선택된 상점가 (null이면 선택 없음)
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
  const defaultCenter = new kakao.maps.LatLng(36.4805, 127.2895); // 보람동 일대 중심
  map = new kakao.maps.Map(container, {
    center: defaultCenter,
    level: 5
  });
  geocoder = new kakao.maps.services.Geocoder();

  // 줌 레벨 변경 시 라벨 표시/숨김 (마커는 항상 유지)
  kakao.maps.event.addListener(map, "zoom_changed", () => {
    updateLabelVisibilityByZoom();
  });

  // 화면 픽셀 좌표(clientX, clientY) -> 지도 위경도 변환 (PC 우클릭 / 모바일 꾹 누르기 공통 사용)
  function containerPointToLatLng(clientX, clientY) {
    const rect = container.getBoundingClientRect();
    const x = clientX - rect.left;
    const y = clientY - rect.top;
    return map.getProjection().coordsFromContainerPoint(new kakao.maps.Point(x, y));
  }

  // 우클릭 직후 브라우저가 보내는 가짜 click 을 무시하기 위한 타임스탬프
  let ignoreClickUntil = 0;

  // 우클릭(PC) - 카카오맵 rightclick 이벤트 대신 표준 contextmenu 이벤트를 직접 사용
  container.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    ignoreClickUntil = Date.now() + 400;
    const latlng = containerPointToLatLng(e.clientX, e.clientY);
    handleMapPick(latlng);
  });

  // 일반 클릭: 이미 핀이 찍힌 영역 안을 클릭하면 제거
  kakao.maps.event.addListener(map, "click", (mouseEvent) => {
    if (Date.now() < ignoreClickUntil) return;
    handleMapClick(mouseEvent.latLng);
  });

  // DOM 클릭 백업: 카카오 오버레이가 이벤트를 삼켜도 좌표로 직접 판별해 제거
  container.addEventListener("click", (e) => {
    if (e.button !== 0) return;
    if (Date.now() < ignoreClickUntil) return;
    const latlng = containerPointToLatLng(e.clientX, e.clientY);
    handleMapClick(latlng);
  });

  // 꾹 누르기(모바일) - 터치 길게 누르기 직접 구현 (PC와 동일한 좌표 변환 함수 사용)
  let longPressTimer = null;
  let longPressStartXY = null;

  container.addEventListener("touchstart", (e) => {
    if (e.touches.length !== 1) return;
    const touch = e.touches[0];
    longPressStartXY = { x: touch.clientX, y: touch.clientY };

    longPressTimer = setTimeout(() => {
      ignoreClickUntil = Date.now() + 500;
      const latlng = containerPointToLatLng(touch.clientX, touch.clientY);
      handleMapPick(latlng);
    }, 550);
  }, { passive: true });

  const cancelLongPress = (e) => {
    if (longPressTimer && longPressStartXY && e.changedTouches && e.changedTouches[0]) {
      const touch = e.changedTouches[0];
      const moved = Math.abs(touch.clientX - longPressStartXY.x) + Math.abs(touch.clientY - longPressStartXY.y);
      if (moved > 12) clearTimeout(longPressTimer);
    } else if (longPressTimer) {
      clearTimeout(longPressTimer);
    }
  };
  container.addEventListener("touchmove", cancelLongPress, { passive: true });
  container.addEventListener("touchend", () => clearTimeout(longPressTimer));
  container.addEventListener("touchcancel", () => clearTimeout(longPressTimer));
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

// 라벨 클릭 3단계 사이클:
// 1번 클릭: 그 구역만 진한 파란색으로 강조 (나머지는 하늘색 유지)
// 2번 클릭(같은 라벨): 그 상점가에 속한 주소 체크리스트(흰색 박스)를 표시
// 3번 클릭(같은 라벨): 완전히 원래 상태로 복귀 (체크리스트 닫힘, 강조 해제)
// 체크리스트의 X 버튼을 누르면 강조는 유지한 채 체크리스트만 닫힘
// 지도의 줌 레벨/범위는 변경하지 않고 색상/체크리스트만 바꿈
function selectMarketByLabel(marketName) {
  if (selectedMarket !== marketName) {
    closeChecklist();
    selectedMarket = marketName;
    applyZoneColorState();
    return;
  }

  if (checklistMarket !== marketName) {
    openChecklist(marketName);
    return;
  }

  closeChecklist();
  selectedMarket = null;
  applyZoneColorState();
}

// 체크리스트에서 켠 주소 강조 폴리곤을 모두 지움
function clearChecklistHighlights() {
  Object.values(checklistHighlights).forEach(polygon => polygon.setMap(null));
  checklistHighlights = {};
}

// 체크리스트(흰색 박스)와 그 안에서 켠 주소 강조를 모두 닫음 (X 버튼 클릭 시에도 호출됨)
function closeChecklist() {
  if (checklistOverlay) {
    checklistOverlay.setMap(null);
    checklistOverlay = null;
  }
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

// 특정 상점가에 속한 주소 체크리스트(흰색 박스)를 라벨 위치에 표시
function openChecklist(marketName) {
  closeChecklist();

  const parcels = MAP_DATA.parcels.filter(p => p.market === marketName);
  const zones = getZonesForMarket(marketName);
  const allCoords = zones.flatMap(z => z.coords);
  const topPoint = getHighestLatPoint(allCoords);
  if (!topPoint) return;

  const position = new kakao.maps.LatLng(topPoint.lat, topPoint.lng);

  const box = document.createElement("div");
  box.className = "addr-checklist";

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
  box.appendChild(closeBtn);

  const title = document.createElement("div");
  title.className = "addr-checklist-title";
  title.textContent = `${getMarketDisplayName(marketName)} 주소`;
  box.appendChild(title);

  if (!parcels.length) {
    const empty = document.createElement("div");
    empty.className = "addr-checklist-empty";
    empty.textContent = "등록된 주소가 없습니다.";
    box.appendChild(empty);
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
    box.appendChild(row);
  });

  checklistOverlay = new kakao.maps.CustomOverlay({
    map,
    position,
    content: box,
    yAnchor: 1,
    xAnchor: 0.5,
    zIndex: 60
  });

  checklistMarket = marketName;
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
    content.addEventListener("click", () => selectMarketByLabel(m.baseName));

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
    renderInitialOverview();
  } else {
    const last = pickedLocations[pickedLocations.length - 1];
    renderPickedLocationDetails(last.latlng, last.jibunFull, last.roadFull, last.parcelAddress, pickedLocations.length);
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
      key: `free|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`,
      showPurple: false
    });
  });
}

// 지도/폴리곤 왼쪽 클릭 → 해당 영역/근처 핀 제거
function handleMapClick(latlng) {
  removePickedAtLatLng(latlng);
}

function addPickedPin({ center, hitCoords, labelText, jibunFull, roadFull, parcelAddress, key, showPurple }) {
  // 좌표를 다시 숫자로 복사해 새 LatLng 생성 (참조/변형 문제 방지)
  const pos = new kakao.maps.LatLng(center.getLat(), center.getLng());
  const id = ++pickedLocationIdSeq;
  const areaKey = getAreaKey(key);

  // 구역/필지일 때만 보라색 표시, 같은 구역은 폴리곤 1개만 유지
  if (showPurple && hitCoords && hitCoords.length) {
    ensureAreaPolygon(areaKey, hitCoords);
  }

  const marker = new kakao.maps.Marker({
    map,
    position: pos,
    zIndex: 50
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
    zIndex: 51,
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
    parcelKey: key
  });

  renderPickedLocationDetails(pos, jibunFull || "", roadFull || "", parcelAddress || "", pickedLocations.length);
}

function renderPickedLocationDetails(latlng, jibunFull, roadFull, parcelAddress, count) {
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");
  const list = document.getElementById("resultList");

  const n = count != null ? count : pickedLocations.length;
  title.textContent = n > 1 ? `선택한 위치 상세정보 (${n}개)` : "선택한 위치 상세정보";
  badge.style.display = "none";

  list.innerHTML = `
    <div class="detail-block">
      <div class="detail-row"><span class="d-label">위치</span><span class="d-value">${parcelAddress || jibunFull || "확인되지 않음"}</span></div>
      <div class="detail-row"><span class="d-label">지번주소</span><span class="d-value">${jibunFull || "확인되지 않음"}</span></div>
      <div class="detail-row"><span class="d-label">도로명주소</span><span class="d-value">${roadFull || "확인되지 않음"}</span></div>
      <div class="detail-row"><span class="d-label">위도</span><span class="d-value">${latlng.getLat().toFixed(6)}</span></div>
      <div class="detail-row"><span class="d-label">경도</span><span class="d-value">${latlng.getLng().toFixed(6)}</span></div>
      <div class="detail-row"><span class="d-label">안내</span><span class="d-value">보라색 영역 안을 클릭(또는 다시 우클릭)하면 핀이 사라집니다.</span></div>
    </div>
  `;
}

/* =========================================================
   7. 결과 패널 렌더링
   ========================================================= */
function renderResultList(parcels) {
  const list = document.getElementById("resultList");
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");
  badge.style.display = "";

  if (!parcels.length) {
    list.innerHTML = `<div class="result-empty">일치하는 결과가 없습니다. 시장명(예: 보람동 호려울) 또는 지번 주소로 검색해보세요.</div>`;
    title.textContent = "검색 결과";
    badge.textContent = "골목형상점가 0곳";
    return;
  }

  title.textContent = `검색 결과 (${parcels.length}건)`;

  const matchedMarkets = [...new Set(parcels.map(p => p.market))];
  const alleyCount = matchedMarkets.filter(m => getMarketType(m) === "골목형상점가").length;
  badge.textContent = `골목형상점가 ${alleyCount}곳`;

  list.innerHTML = parcels.map(p => `
    <div class="result-item" data-market="${p.market}" data-id="${p.id}">
      <span class="r-market">${getMarketDisplayName(p.market)}</span>
      <span class="r-addr">${formatParcelAddress(p)}</span>
    </div>
  `).join("");

  list.querySelectorAll(".result-item").forEach(el => {
    el.addEventListener("click", () => {
      const id = Number(el.dataset.id);
      const parcel = MAP_DATA.parcels.find(p => p.id === id);
      if (parcel) focusOnParcels([parcel]);
    });
  });
}

// 초기 로딩 시: 결과 패널에 전체 골목형상점가 목록을 보여줌 (지도에는 이미 구역 배경이 항상 표시되어 있음)
function renderInitialOverview() {
  const alleyMarkets = MAP_DATA.markets.filter(m => m.type === "골목형상점가");

  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");
  const list = document.getElementById("resultList");
  badge.style.display = "";

  title.textContent = "골목형상점가 전체 구역도";
  badge.textContent = `골목형상점가 ${alleyMarkets.length}곳`;

  if (!alleyMarkets.length) {
    list.innerHTML = `<div class="result-empty">표시할 골목형상점가 데이터가 없습니다.</div>`;
    return;
  }

  list.innerHTML = alleyMarkets.map(m => {
    const zoneCount = getZonesForMarket(m.baseName).length;
    return `
      <div class="result-item" data-market="${m.baseName}">
        <span class="r-market">${m.name}</span>
        <span class="r-addr">구역 ${zoneCount}개</span>
      </div>
    `;
  }).join("");

  list.querySelectorAll(".result-item").forEach(el => {
    el.addEventListener("click", () => {
      selectMarketByLabel(el.dataset.market);
    });
  });
}

/* =========================================================
   8. 검색 처리
   ========================================================= */
function geocodeFallbackSearch(query) {
  if (!geocoder) return;
  geocoder.addressSearch(query, (result, status) => {
    if (status === kakao.maps.services.Status.OK && result[0]) {
      const coords = new kakao.maps.LatLng(result[0].y, result[0].x);

      clearHighlights();
      closeChecklist();
      selectedMarket = null;
      applyZoneColorState();

      map.setCenter(coords);
      map.setLevel(4);
      const marker = new kakao.maps.Marker({ map, position: coords });
      highlightOverlays.push({ marker });

      const list = document.getElementById("resultList");
      list.innerHTML = `<div class="result-empty">등록된 상점가 데이터에는 없는 주소입니다. 입력하신 주소 위치로 지도를 이동했습니다.</div>`;
      document.getElementById("resultTitle").textContent = "검색 결과";
      document.getElementById("alleyCountBadge").style.display = "";
      document.getElementById("alleyCountBadge").textContent = "골목형상점가 0곳";
    } else {
      const list = document.getElementById("resultList");
      list.innerHTML = `<div class="result-empty">검색 결과가 없습니다. 시장명 또는 정확한 주소를 입력해주세요.</div>`;
      document.getElementById("resultTitle").textContent = "검색 결과";
      document.getElementById("alleyCountBadge").style.display = "";
      document.getElementById("alleyCountBadge").textContent = "골목형상점가 0곳";
    }
  });
}

function handleSearch() {
  const query = document.getElementById("searchInput").value;
  if (!query.trim()) return;

  const matched = searchParcels(query);
  renderResultList(matched);

  if (matched.length) {
    focusOnParcels(matched);
  } else if (map) {
    geocodeFallbackSearch(query);
  }
}

document.getElementById("searchBtn").addEventListener("click", handleSearch);
document.getElementById("searchInput").addEventListener("keydown", (e) => {
  if (e.key === "Enter") handleSearch();
});

/* -------- 초기화 버튼: 사이트 첫 진입 상태로 복귀 -------- */
function resetToInitialView() {
  document.getElementById("searchInput").value = "";
  clearHighlights();
  clearPickedLocation();
  closeChecklist();
  selectedMarket = null;

  activeTypes = new Set(Object.keys(TYPE_COLORS));
  document.querySelectorAll(".type-chip").forEach(chip => chip.classList.add("active"));
  applyTypeVisibility();

  applyZoneColorState();
  renderInitialOverview();

  if (MAP_DATA.zones.length) {
    const allPaths = MAP_DATA.zones.map(z => toLatLngPath(z.coords));
    fitBoundsToPaths(allPaths);
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
    const allPaths = MAP_DATA.zones.map(z => toLatLngPath(z.coords));
    fitBoundsToPaths(allPaths);
  });
