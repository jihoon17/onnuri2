/* =========================================================
   상점가 위치 검색 지도 - 기능 로직
   ========================================================= */

/* -------- 설정 -------- */
// 카카오 개발자센터에서 발급받은 JavaScript 키
const KAKAO_APP_KEY = "01a0f8272692845e09fe8d4c402e6317";

// 엑셀 원본 파일 경로 (이 index.html과 같은 폴더/레포에 올려주세요)
const EXCEL_FILE_URL = "mapData_v1.xlsx";

// 색상 (style.css의 CSS 변수와 값 맞춰둠)
const COLOR_ZONE_DEFAULT = "#3b6ea5";  // 구역 기본 색 (파란색)
const COLOR_ZONE_SELECTED = "#123a63"; // 라벨 클릭 시 선택된 구역 (진한 파란색)
const COLOR_ZONE_DIMMED = "#8fd0f2";   // 라벨 클릭 시 선택 외 구역 (하늘색)
const COLOR_HIGHLIGHT = "#e02424";     // 검색된 특정 주소 (빨간색)

/* -------- 전역 상태 -------- */
let map, geocoder;
let MAP_DATA = { markets: [], parcels: [], zones: [] };

let zoneOverlaysByMarket = {};  // marketName -> [kakao.maps.Polygon, ...] (항상 유지되는 구역 배경)
let marketLabelOverlays = [];   // { marketName, marker, overlay, content } (항상 유지되는 라벨)
let highlightOverlays = [];     // { polygon, marker } (검색 시에만 갱신/삭제되는 강조 표시)

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
function applyZoneColorState() {
  Object.entries(zoneOverlaysByMarket).forEach(([marketName, polygons]) => {
    let color, fillOpacity;
    if (!selectedMarket) {
      color = COLOR_ZONE_DEFAULT;
      fillOpacity = 0.4;
    } else if (marketName === selectedMarket) {
      color = COLOR_ZONE_SELECTED;
      fillOpacity = 0.55;
    } else {
      color = COLOR_ZONE_DIMMED;
      fillOpacity = 0.35;
    }
    polygons.forEach(p => p.setOptions({ strokeColor: color, fillColor: color, fillOpacity }));
  });

  marketLabelOverlays.forEach(o => {
    o.content.classList.toggle("selected", o.marketName === selectedMarket);
  });
}

// 라벨 클릭: 클릭한 상점가 구역은 진한 파란색, 나머지는 하늘색. 같은 라벨 다시 클릭 시 기본 상태로 복귀
function selectMarketByLabel(marketName) {
  selectedMarket = (selectedMarket === marketName) ? null : marketName;
  applyZoneColorState();

  if (selectedMarket) {
    const paths = getZonesForMarket(selectedMarket).map(z => toLatLngPath(z.coords));
    fitBoundsToPaths(paths);
  } else {
    const paths = MAP_DATA.zones.map(z => toLatLngPath(z.coords));
    fitBoundsToPaths(paths);
  }
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
   7. 결과 패널 렌더링
   ========================================================= */
function renderResultList(parcels) {
  const list = document.getElementById("resultList");
  const title = document.getElementById("resultTitle");
  const badge = document.getElementById("alleyCountBadge");

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
      document.getElementById("alleyCountBadge").textContent = "골목형상점가 0곳";
    } else {
      const list = document.getElementById("resultList");
      list.innerHTML = `<div class="result-empty">검색 결과가 없습니다. 시장명 또는 정확한 주소를 입력해주세요.</div>`;
      document.getElementById("resultTitle").textContent = "검색 결과";
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
