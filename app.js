/* =========================================================
   상점가 위치 검색 지도 - 기능 로직
   ========================================================= */

/* -------- 설정 -------- */
// 카카오 개발자센터에서 발급받은 JavaScript 키
const KAKAO_APP_KEY = "01a0f8272692845e09fe8d4c402e6317";

// 엑셀 원본 파일 경로 (이 index.html과 같은 폴더/레포에 올려주세요)
const EXCEL_FILE_URL = "mapData_v1.xlsx";

// 색상 (style.css의 CSS 변수와 값 맞춰둠)
const COLOR_ZONE_DEFAULT = "#8fd0f2";  // 구역 기본 색 (하늘색) - 항상 이 색이 기본
const COLOR_ZONE_SELECTED = "#123a63"; // 라벨 클릭 시 선택된 구역만 진하게 (진한 파란색)
const COLOR_HIGHLIGHT = "#e02424";     // 검색된 특정 주소 (빨간색)

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

/* =========================================================
   1. 엑셀 원본 읽기
   ========================================================= */
async function loadExcelData(url) {
  const resp = await fetch(url);
  if (!resp.ok) throw new Error("EXCEL_FETCH_FAIL");
  const buf = await resp.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const overviewRows = XLSX.utils.sheet_to_json(wb.Sheets["시장개요"], { defval: "" });
  const addrRows = XLSX.utils.sheet_to_json(wb.Sheets["시장별주소"], { defval: "" });
  const zoneRows = XLSX.utils.sheet_to_json(wb.Sheets["구역"], { defval: "" });

  const markets = overviewRows.map(r => ({
    id: r["순번"],
    type: String(r["구분"]).trim(),
    name: String(r["명칭"]).trim()
  }));

  const parcels = addrRows.map(r => ({
    id: r["순번"],
    market: String(r["소속시장"]).trim(),
    address: String(r["주소"]).trim(),
    coords: JSON.parse(r["경계좌표"])
  }));

  const zones = zoneRows.map(r => ({
    id: r["순번"],
    market: String(r["소속시장"]).trim(),
    zoneNo: r["구역순번"],
    coords: JSON.parse(r["경계좌표"])
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
  const m = MAP_DATA.markets.find(m => m.name === marketName);
  return m ? m.type : null;
}

// 좌표 목록 중 위도(lat)가 가장 높은 좌표 반환
function getHighestLatPoint(coordsList) {
  let best = null;
  coordsList.forEach(c => {
    if (!best || c.lat > best.lat) best = c;
  });
  return best;
}

// 검색어와 매칭되는 필지(주소) 목록 (시장명 또는 주소 포함검색)
function searchParcels(query) {
  const q = query.trim();
  if (!q) return [];
  return MAP_DATA.parcels.filter(p =>
    p.market.includes(q) || p.address.includes(q) || q.includes(p.market)
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
    const zones = getZonesForMarket(m.name);
    const polygons = zones.map(z => {
      const path = toLatLngPath(z.coords);
      return new kakao.maps.Polygon({
        map,
        path,
        strokeWeight: 2,
        strokeColor: COLOR_ZONE_DEFAULT,
        strokeOpacity: 0.9,
        fillColor: COLOR_ZONE_DEFAULT,
        fillOpacity: 0.4,
        zIndex: 1
      });
    });
    zoneOverlaysByMarket[m.name] = polygons;
  });
}

// 현재 selectedMarket 상태에 맞춰 구역 색상을 다시 칠함 (다시 그리지 않고 옵션만 변경)
// 기본 색은 항상 하늘색이며, 선택된 구역만 진한 파란색으로 강조. 선택 안 된 나머지 구역은
// 옅어지지 않고 기본(하늘색) 그대로 유지됨.
function applyZoneColorState() {
  Object.entries(zoneOverlaysByMarket).forEach(([marketName, polygons]) => {
    let color, fillOpacity;
    if (marketName === selectedMarket) {
      color = COLOR_ZONE_SELECTED;
      fillOpacity = 0.55;
    } else {
      color = COLOR_ZONE_DEFAULT;
      fillOpacity = 0.4;
    }
    polygons.forEach(p => p.setOptions({ strokeColor: color, fillColor: color, fillOpacity }));
  });

  marketLabelOverlays.forEach(o => {
    o.content.classList.toggle("selected", o.marketName === selectedMarket);
  });
}

// 라벨 클릭: 클릭한 상점가 구역은 진한 파란색, 나머지는 하늘색. 같은 라벨 다시 클릭 시 기본 상태로 복귀
// 지도의 줌 레벨/범위는 변경하지 않고 색상만 바꿈
function selectMarketByLabel(marketName) {
  selectedMarket = (selectedMarket === marketName) ? null : marketName;
  applyZoneColorState();
}

/* =========================================================
   5. 상점가 명칭 라벨 레이어 - 각 구역 중 위도가 가장 높은 지점에 표시
   ========================================================= */
function drawMarketLabels() {
  marketLabelOverlays.forEach(o => {
    if (o.marker) o.marker.setMap(null);
    if (o.overlay) o.overlay.setMap(null);
  });
  marketLabelOverlays = [];

  MAP_DATA.markets.forEach(m => {
    const zones = getZonesForMarket(m.name);
    if (!zones.length) return;

    const allCoords = zones.flatMap(z => z.coords);
    const topPoint = getHighestLatPoint(allCoords);
    if (!topPoint) return;

    const position = new kakao.maps.LatLng(topPoint.lat, topPoint.lng);

    const marker = new kakao.maps.Marker({
      map,
      position,
      zIndex: 20
    });

    const content = document.createElement("div");
    content.className = "market-label";
    content.textContent = m.name;
    content.addEventListener("click", () => selectMarketByLabel(m.name));

    const overlay = new kakao.maps.CustomOverlay({
      map,
      position,
      content,
      yAnchor: 1,
      xAnchor: 0.5,
      zIndex: 21
    });

    marketLabelOverlays.push({ marketName: m.name, marker, overlay, content });
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

  // 라벨 선택 상태를 초기화해서 모든 구역이 다시 기본 파란색으로 보이게 함
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

  // 1) 필지 안 → 우클릭한 그 위치에 마커 + 필지 보라색
  const parcel = findParcelAt(clickPos);
  if (parcel) {
    addPickedPin({
      center: clickPos,
      hitCoords: parcel.coords,
      labelText: parcel.address || parcel.market || "선택한 위치",
      jibunFull: parcel.address || "",
      roadFull: "",
      parcelAddress: parcel.address || "",
      key: `parcel|${parcel.market}|${parcel.address}|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`
    });
    return;
  }

  // 2) 구역 안 → 우클릭한 그 위치에 마커 + 구역 보라색
  const zone = findZoneAt(clickPos);
  if (zone) {
    addPickedPin({
      center: clickPos,
      hitCoords: zone.coords,
      labelText: zone.market ? `${zone.market} 구역${zone.zoneNo || ""}` : "선택한 구역",
      jibunFull: "",
      roadFull: "",
      parcelAddress: zone.market || "",
      key: `zone|${zone.market}|${zone.zoneNo}|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`
    });
    return;
  }

  // 3) 그 외 → 클릭 지점 + 원형 보라색(~35m)
  const circle = makeCircleCoords(clickLat, clickLng, 35);

  if (!geocoder) {
    addPickedPin({
      center: clickPos,
      hitCoords: circle,
      labelText: "선택한 위치",
      jibunFull: "",
      roadFull: "",
      parcelAddress: "",
      key: `free|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`
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
    // 비동기 콜백에서도 동일한 clickPos 사용 (중앙으로 끌어가지 않음)
    addPickedPin({
      center: clickPos,
      hitCoords: circle,
      labelText,
      jibunFull,
      roadFull,
      parcelAddress: jibunFull,
      key: `free|${clickLat.toFixed(6)}|${clickLng.toFixed(6)}`
    });
  });
}

// 지도/폴리곤 왼쪽 클릭 → 해당 영역/근처 핀 제거
function handleMapClick(latlng) {
  removePickedAtLatLng(latlng);
}

function addPickedPin({ center, hitCoords, labelText, jibunFull, roadFull, parcelAddress, key }) {
  // 좌표를 다시 숫자로 복사해 새 LatLng 생성 (참조/변형 문제 방지)
  const pos = new kakao.maps.LatLng(center.getLat(), center.getLng());
  const id = ++pickedLocationIdSeq;
  const areaKey = getAreaKey(key);

  // 같은 구역/필지는 보라색 폴리곤을 하나만 유지 (겹쳐서 진해지지 않음)
  ensureAreaPolygon(areaKey, hitCoords);

  const marker = new kakao.maps.Marker({
    map,
    position: pos,
    zIndex: 50
  });

  const content = document.createElement("div");
  content.className = "market-label picked-label";
  content.textContent = labelText;
  content.title = "마커 또는 보라색 영역을 클릭하면 사라집니다";
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
    areaKey,
    hitCoords,
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
      <span class="r-market">${p.market}</span>
      <span class="r-addr">${p.address}</span>
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
    const zoneCount = getZonesForMarket(m.name).length;
    return `
      <div class="result-item" data-market="${m.name}">
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
  selectedMarket = null;
  applyZoneColorState();
  renderInitialOverview();

  if (MAP_DATA.zones.length) {
    const allPaths = MAP_DATA.zones.map(z => toLatLngPath(z.coords));
    fitBoundsToPaths(allPaths);
  }
}

document.getElementById("resetBtn").addEventListener("click", resetToInitialView);

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
Promise.all([loadKakaoSdk(KAKAO_APP_KEY), loadExcelData(EXCEL_FILE_URL)])
  .then(([_, data]) => {
    MAP_DATA = data;
    initMap();
    drawAllZonesBase();   // 모든 구역을 파란색으로 항상 표시
    drawMarketLabels();   // 각 구역 최고 위도 지점에 명칭 라벨 표시
    renderInitialOverview();
    const allPaths = MAP_DATA.zones.map(z => toLatLngPath(z.coords));
    fitBoundsToPaths(allPaths);
  })
  .catch((err) => {
    if (err.message === "NO_KEY") {
      showFallback("카카오 지도 API 키가 설정되지 않았습니다.");
    } else if (err.message === "EXCEL_FETCH_FAIL") {
      showFallback(`엑셀 데이터 파일(${EXCEL_FILE_URL})을 불러오지 못했습니다. index.html과 같은 위치에 파일이 있는지 확인해주세요.`);
    } else {
      showFallback("카카오 지도를 불러오지 못했습니다.");
    }
  });
