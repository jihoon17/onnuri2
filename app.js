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
let pickedLocations = [];       // [{ id, marker, overlay }, ...] (우클릭/꾹 누르기로 찍은 여러 위치 핀)

let selectedMarket = null;      // 라벨 클릭으로 선택된 상점가 (null이면 선택 없음)
let checklistMarket = null;     // 체크리스트가 열려있는 상점가 (null이면 닫힘)
let checklistOverlay = null;    // 체크리스트 흰색 박스(CustomOverlay)
let checklistHighlights = {};   // parcelId -> kakao.maps.Polygon (체크리스트에서 켠 주소 강조)

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

  // 우클릭(PC) - 카카오맵 rightclick 이벤트 대신 표준 contextmenu 이벤트를 직접 사용
  // (카카오맵 SDK의 rightclick 이벤트가 환경에 따라 발생하지 않는 경우가 있어 더 안정적인 방식으로 통일)
  container.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    const latlng = containerPointToLatLng(e.clientX, e.clientY);
    handleMapPick(latlng);
  });

  // 꾹 누르기(모바일) - 터치 길게 누르기 직접 구현 (PC와 동일한 좌표 변환 함수 사용)
  let longPressTimer = null;
  let longPressStartXY = null;

  container.addEventListener("touchstart", (e) => {
    if (e.touches.length !== 1) return;
    const touch = e.touches[0];
    longPressStartXY = { x: touch.clientX, y: touch.clientY };

    longPressTimer = setTimeout(() => {
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

// 라벨 클릭 3단계 사이클:
// 1번 클릭: 그 구역만 진한 파란색으로 강조 (나머지는 하늘색 유지)
// 2번 클릭(같은 라벨): 그 상점가에 속한 주소 체크리스트(흰색 박스)를 표시
// 3번 클릭(같은 라벨): 완전히 원래 상태로 복귀 (체크리스트 닫힘, 강조 해제)
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

// 체크리스트(흰색 박스)와 그 안에서 켠 주소 강조를 모두 닫음
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

  const title = document.createElement("div");
  title.className = "addr-checklist-title";
  title.textContent = `${marketName} 주소`;
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
    span.textContent = p.address;

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
    zIndex: 40
  });

  checklistMarket = marketName;
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

  // 라벨 선택/체크리스트 상태를 초기화해서 모든 구역이 다시 기본 색으로 보이게 함
  closeChecklist();
  selectedMarket = null;
  applyZoneColorState();

  const paths = parcels.map(p => drawHighlightPolygon(p));
  fitBoundsToPaths(paths);
}

/* =========================================================
   6-1. 우클릭(PC) / 꾹 누르기(모바일) - 임의 위치 핀 찍기
   (여러 개를 동시에 찍을 수 있고, 같은 마커/라벨을 다시 클릭하면 그 핀만 사라짐)
   ========================================================= */
let pickedLocationSeq = 0;

function clearAllPickedLocations() {
  pickedLocations.forEach(p => {
    p.marker.setMap(null);
    p.overlay.setMap(null);
  });
  pickedLocations = [];
}

function removePickedLocation(id) {
  const idx = pickedLocations.findIndex(p => p.id === id);
  if (idx === -1) return;
  pickedLocations[idx].marker.setMap(null);
  pickedLocations[idx].overlay.setMap(null);
  pickedLocations.splice(idx, 1);
}

// 전체 주소에서 맨 앞의 시/도 이름만 제거 (예: "세종특별자치시 보람동 721" -> "보람동 721")
function stripCityName(fullAddress, cityName) {
  if (!fullAddress) return "";
  if (cityName && fullAddress.startsWith(cityName)) {
    return fullAddress.slice(cityName.length).trim();
  }
  return fullAddress;
}

function handleMapPick(latlng) {
  if (!geocoder) return;

  geocoder.coord2Address(latlng.getLng(), latlng.getLat(), (result, status) => {
    if (status === kakao.maps.services.Status.OK && result[0]) {
      addPickedLocation(latlng, result[0]);
    } else {
      addPickedLocation(latlng, null);
    }
  });
}

function addPickedLocation(latlng, addrResult) {
  const id = ++pickedLocationSeq;

  let jibunFull = "";
  let roadFull = "";
  let labelText = "선택한 위치";

  if (addrResult) {
    if (addrResult.address) {
      jibunFull = addrResult.address.address_name;
      labelText = stripCityName(jibunFull, addrResult.address.region_1depth_name) || jibunFull;
    }
    if (addrResult.road_address) {
      roadFull = addrResult.road_address.address_name;
    }
  }

  const marker = new kakao.maps.Marker({
    map,
    position: latlng,
    zIndex: 30,
    clickable: true
  });

  const content = document.createElement("div");
  content.className = "market-label picked-label";
  content.textContent = labelText;
  content.title = "클릭하면 이 마커가 사라집니다";
  content.addEventListener("click", () => removePickedLocation(id));

  const overlay = new kakao.maps.CustomOverlay({
    map,
    position: latlng,
    content,
    yAnchor: 1,
    xAnchor: 0.5,
    zIndex: 31
  });

  kakao.maps.event.addListener(marker, "click", () => removePickedLocation(id));

  pickedLocations.push({ id, marker, overlay });

  renderPickedLocationDetails(latlng, jibunFull, roadFull);
}

function renderPickedLocationDetails(latlng, jibunFull, roadFull) {
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");
  const list = document.getElementById("resultList");

  title.textContent = "선택한 위치 상세정보";
  badge.style.display = "none";

  list.innerHTML = `
    <div class="detail-block">
      <div class="detail-row"><span class="d-label">지번주소</span><span class="d-value">${jibunFull || "확인되지 않음"}</span></div>
      <div class="detail-row"><span class="d-label">도로명주소</span><span class="d-value">${roadFull || "확인되지 않음"}</span></div>
      <div class="detail-row"><span class="d-label">위도</span><span class="d-value">${latlng.getLat().toFixed(6)}</span></div>
      <div class="detail-row"><span class="d-label">경도</span><span class="d-value">${latlng.getLng().toFixed(6)}</span></div>
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
  clearAllPickedLocations();
  closeChecklist();
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
