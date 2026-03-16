// ===== GLOBALS =====
const items = DATA.items;
const colors = DATA.colors;
const stats = DATA.stats;

let activeRegionals = new Set();
let polygonLayers = [];
let searchTerm = '';
let somenteVinculados = false;

// ===== HELPERS =====
const fmtNum = (n) => n ? n.toLocaleString('pt-BR') : '-';
const fmtBRL = (n) => {
  if (!n) return '-';
  if (Math.abs(n) >= 1000) {
    const bi = n / 1000;
    return 'R$ ' + bi.toLocaleString('pt-BR', {minimumFractionDigits: bi % 1 === 0 ? 0 : 2, maximumFractionDigits: 2}) + ' bi';
  }
  return 'R$ ' + n.toLocaleString('pt-BR', {minimumFractionDigits: n % 1 === 0 ? 0 : 1, maximumFractionDigits: 1}) + ' mi';
};
const fmtArea = (n) => n ? (n / 10000).toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 0}) + ' ha' : '-';

function isLinked(item) {
  return !!(item.e && item.e.regional !== null && item.e.regional !== undefined);
}

function getCentroid(item) {
  if (item.c) return item.c;
  if (item.p && item.p.length > 0 && item.p[0].length > 0) {
    const coords = item.p[0];
    const lat = coords.reduce((s, c) => s + c[0], 0) / coords.length;
    const lng = coords.reduce((s, c) => s + c[1], 0) / coords.length;
    return [lat, lng];
  }
  return null;
}

// ===== STATS =====
document.getElementById('statsGrid').innerHTML = `
  <div class="stat-card stat-card-full">
    <div class="val">230</div>
    <div class="label">Empreendimentos</div>
  </div>
  <div class="stat-card">
    <div class="val">${fmtNum(Math.round(stats.total_units))}</div>
    <div class="label">Total Unidades</div>
  </div>
  <div class="stat-card">
    <div class="val">${fmtArea(stats.total_area)}</div>
    <div class="label">Área Total</div>
  </div>
  <div class="stat-card">
    <div class="val green">${fmtBRL(stats.total_vgv)}</div>
    <div class="label">VGV Total</div>
  </div>
  <div class="stat-card">
    <div class="val green">${fmtBRL(stats.total_vgv_bt)}</div>
    <div class="label">VGV Total BT</div>
  </div>
`;

// ===== MAP INIT =====
const map = L.map('map', {zoomControl: false, attributionControl: false}).setView([-12, -50], 4);
L.control.zoom({position: 'topright'}).addTo(map);

// ===== TILE LAYERS =====
const streetLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
  attribution: '© OpenStreetMap', maxZoom: 19
}).addTo(map);

const satelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
  attribution: '© Esri', maxZoom: 19
});

const terrainLayer = L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', {
  attribution: '© OpenTopoMap', maxZoom: 17, opacity: 0.7
});

const vegetationLayer = L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', {
  attribution: '© OpenTopoMap', maxZoom: 17, opacity: 0.55
});

const layerMap = {
  layerStreets: streetLayer,
  layerSatellite: satelliteLayer,
  layerTerrain: terrainLayer,
  layerVegetation: vegetationLayer
};

// ===== LAYER GROUPS =====
const ZOOM_THRESHOLD = 12;
const layerGroup    = L.layerGroup().addTo(map);
const overviewGroup = L.layerGroup().addTo(map);

// ===== LAYER TOGGLE LOGIC =====
function handleLayerToggle(id) {
  const cb = document.getElementById(id);
  const layer = layerMap[id];
  if (cb.checked) {
    if (id === 'layerStreets') {
      document.getElementById('layerSatellite').checked = false;
      map.removeLayer(satelliteLayer);
    } else if (id === 'layerSatellite') {
      document.getElementById('layerStreets').checked = false;
      map.removeLayer(streetLayer);
    }
    map.addLayer(layer);
    if (id === 'layerStreets' || id === 'layerSatellite') layer.bringToBack();
  } else {
    map.removeLayer(layer);
    if (id === 'layerStreets' && !document.getElementById('layerSatellite').checked) {
      cb.checked = true;
      return;
    }
    if (id === 'layerSatellite' && !document.getElementById('layerStreets').checked) {
      document.getElementById('layerStreets').checked = true;
      map.addLayer(streetLayer);
      streetLayer.bringToBack();
    }
  }
}

Object.keys(layerMap).forEach(id => {
  document.getElementById(id).addEventListener('change', () => handleLayerToggle(id));
});

// Layer panel collapse
let panelOpen = true;
document.getElementById('layerToggle').addEventListener('click', () => {
  panelOpen = !panelOpen;
  document.getElementById('layerBody').classList.toggle('collapsed', !panelOpen);
  document.getElementById('chevron').innerHTML = panelOpen ? '&#9660;' : '&#9654;';
});

// ===== MOBILE SIDEBAR DRAWER =====
const sidebar       = document.getElementById('sidebar');
const overlay       = document.getElementById('sidebarOverlay');
const mobileMenuBtn = document.getElementById('mobileMenuBtn');
const sidebarCloseBtn = document.getElementById('sidebarClose');

function openSidebar() {
  sidebar.classList.add('open');
  overlay.classList.add('visible');
  document.body.style.overflow = 'hidden';
}
function closeSidebar() {
  sidebar.classList.remove('open');
  overlay.classList.remove('visible');
  document.body.style.overflow = '';
}

if (mobileMenuBtn)  mobileMenuBtn.addEventListener('click', openSidebar);
if (overlay)        overlay.addEventListener('click', closeSidebar);
if (sidebarCloseBtn) sidebarCloseBtn.addEventListener('click', closeSidebar);

// ===== FILTER CHIPS =====
const allRegionals = [...new Set(
  items.filter(i => isLinked(i)).map(i => i.e.regional).filter(Boolean).filter(r => r !== 'None')
)].sort();

const chipsEl = document.getElementById('filterChips');

const allChip = document.createElement('div');
allChip.className = 'chip active';
allChip.textContent = 'Todos';
allChip.onclick = () => {
  activeRegionals.clear();
  document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
  allChip.classList.add('active');
  updateMap();
};
chipsEl.appendChild(allChip);

/* 
const vinculadosChip = document.createElement('div');
vinculadosChip.className = 'chip';
vinculadosChip.innerHTML = `<span class="dot" style="background:#27a06a"></span>Vinculados`;
vinculadosChip.onclick = () => {
  somenteVinculados = !somenteVinculados;
  vinculadosChip.classList.toggle('active', somenteVinculados);
  updateMap();
};
chipsEl.appendChild(vinculadosChip);
*/

allRegionals.forEach(r => {
  const chip = document.createElement('div');
  chip.className = 'chip';
  chip.innerHTML = `<span class="dot" style="background:${colors[r] || '#7f8c8d'}"></span>${r}`;
  chip.onclick = () => {
    if (activeRegionals.has(r)) {
      activeRegionals.delete(r);
      chip.classList.remove('active');
    } else {
      activeRegionals.add(r);
      chip.classList.add('active');
    }
    allChip.classList.toggle('active', activeRegionals.size === 0);
    updateMap();
  };
  chipsEl.appendChild(chip);
});

// ===== SEARCH =====
document.getElementById('searchInput').addEventListener('input', (e) => {
  searchTerm = e.target.value.toUpperCase();
  updateMap();
});

// ===== POPUP CONTENT =====
function popupContent(item) {
  if (!isLinked(item)) {
    return `
      <div class="popup-header">
        <div class="popup-title">${item.n}</div>
      </div>
      <div class="popup-nodata">Área KML sem dados da planilha vinculados.</div>`;
  }

  const e = item.e;
  const regionalColor = colors[e.regional] || '#7f8c8d';
  const statusClass   = e.on_off === 1 ? 'on' : 'off';
  const statusLabel   = e.on_off === 1 ? 'ON' : 'OFF';

  return `
    <div class="popup-header">
      ${e.regional ? `<div class="popup-regional-badge" style="background:${regionalColor}">${e.regional}</div>` : ''}
      ${e.cidade ? `<div class="popup-city">${e.cidade}</div>` : ''}
      <div class="popup-title">Empreendimento:<br> ${e.empreendimento || e.nome || item.n}</div>
    </div>
    <div class="popup-body">
      <div class="popup-grid">
        <div class="popup-cell">
          <div class="pg-label">Tipo</div>
          <div class="pg-val">${e.tipo || '—'}</div>
        </div>
        <div class="popup-cell">
          <div class="pg-label">Ano Prev.</div>
          <div class="pg-val">${e.year || '—'}</div>
        </div>
        <div class="popup-cell">
          <div class="pg-label">Área Total</div>
          <div class="pg-val">${fmtArea(e.area_total)}</div>
        </div>
        <div class="popup-cell">
          <div class="pg-label">Unidades</div>
          <div class="pg-val">${fmtNum(e.total_unidades)}</div>
        </div>

        <div class="popup-divider"></div>

        <div class="popup-cell">
          <div class="pg-label">VGV Total</div>
          <div class="pg-val green">${fmtBRL(e.vgv_total)}</div>
        </div>
        <div class="popup-cell">
          <div class="pg-label">VGV BT</div>
          <div class="pg-val green">${fmtBRL(e.vgv_bt)}</div>
        </div>
        <div class="popup-cell">
          <div class="pg-label">Custo Terreno</div>
          <div class="pg-val">${fmtBRL(e.custo_terreno)}</div>
        </div>
        <div class="popup-cell">
          <div class="pg-label">Custo Construção</div>
          <div class="pg-val">${fmtBRL(e.custo_construcao)}</div>
        </div>

        <div class="popup-divider"></div>

        <div class="popup-cell">
          <div class="pg-label">Part. Buriti</div>
          <div class="pg-val">${e.participacao_buriti ? (e.participacao_buriti * 100).toFixed(1) + '%' : '—'}</div>
        </div>
        <div class="popup-cell">
          <div class="pg-label">Status</div>
          <div class="pg-val"><span class="popup-status ${statusClass}">${statusLabel}</span></div>
        </div>
      </div>
    </div>`;
}

// ===== FILTER LOGIC =====
function passesFilter(item) {
  if (somenteVinculados && !isLinked(item)) return false;
  if (activeRegionals.size > 0) {
    const r = isLinked(item) ? item.e.regional : null;
    if (!r || !activeRegionals.has(r)) return false;
  }
  if (searchTerm) {
    const haystack = [
      item.n,
      item.e ? item.e.nome : '',
      item.e ? item.e.cidade : '',
      item.e ? item.e.empreendimento : '',
      item.e ? item.e.regional : ''
    ].join(' ').toUpperCase();
    if (!haystack.includes(searchTerm)) return false;
  }
  return true;
}

function getColor(item) {
  if (isLinked(item)) return colors[item.e.regional] || '#7f8c8d';
  return '#5a6e8e';
}

// ===== OVERVIEW MARKERS =====
function makeOverviewIcon(color) {
  return L.divIcon({
    className: '',
    html: `<div class="ov-pin" style="--pin-color:${color}"></div>`,
    iconSize:   [22, 30],
    iconAnchor: [11, 30],
    tooltipAnchor: [0, -32],
  });
}

function buildOverviewMarkers() {
  overviewGroup.clearLayers();

  items.forEach((item) => {
    if (!passesFilter(item)) return;
    const centroid = getCentroid(item);
    if (!centroid) return;

    const color    = getColor(item);
    const linked   = isLinked(item);
    const name     = linked ? (item.e.empreendimento || item.e.nome || item.n) : item.n;
    const city     = linked ? (item.e.cidade   || '') : '';
    const regional = linked ? (item.e.regional || '') : '';
    const units    = linked && item.e.total_unidades
      ? fmtNum(item.e.total_unidades) + ' unidades'
      : '';

    const marker = L.marker(centroid, {
      icon: makeOverviewIcon(color),
      zIndexOffset: 200,
    });

    marker.bindTooltip(`
      <div class="ov-tooltip">
        ${regional ? `<span class="ov-tag" style="background:${color}">${regional}</span>` : ''}
        <strong>${name}</strong>
        ${city  ? `<div class="ov-city">${city}</div>`   : ''}
        ${units ? `<div class="ov-units">${units}</div>` : ''}
        <div class="ov-hint">Clique para aproximar</div>
      </div>`, {
      direction: 'top',
      className: 'ov-tooltip-outer',
      offset: [0, -4],
    });

    marker.on('click', () => {
      map.flyTo(centroid, 13, { duration: 1.2, easeLinearity: 0.35 });
    });

    overviewGroup.addLayer(marker);
  });
}

function applyZoomVisibility() {
  const isOverview = map.getZoom() < ZOOM_THRESHOLD;

  overviewGroup.eachLayer(l => {
    const el = l.getElement ? l.getElement() : null;
    if (el) el.style.display = isOverview ? '' : 'none';
  });

  layerGroup.eachLayer(l => {
    const el = l.getElement ? l.getElement() : null;
    if (el) el.style.display = isOverview ? 'none' : '';
  });
}

map.on('zoomend', applyZoomVisibility);

// ===== RENDER MAP + LIST =====
function updateMap() {
  layerGroup.clearLayers();
  polygonLayers = [];
  const listEl = document.getElementById('listContainer');
  listEl.innerHTML = '';
  let visibleCount = 0;

  items.forEach((item, idx) => {
    if (!passesFilter(item)) return;
    visibleCount++;
    const color    = getColor(item);
    const centroid = getCentroid(item);

    // Polígonos
    item.p.forEach(polyCoords => {
      const polygon = L.polygon(polyCoords, {
        color: color, weight: 2, opacity: 0.85,
        fillColor: color, fillOpacity: 0.15, smoothFactor: 1.2
      });
      polygon.on('mouseover', function () {
        this.setStyle({ weight: 3, fillOpacity: 0.28 });
      });
      polygon.on('mouseout', function () {
        this.setStyle({ weight: 2, fillOpacity: 0.15 });
      });
      polygon.bindPopup(popupContent(item), { maxWidth: 320, className: '' });
      polygon.addTo(layerGroup);
      polygonLayers.push({ layer: polygon, item: item, idx: idx, centroid: centroid });
    });

    // Marcador de centroide para itens sem polígono
    if (item.p.length === 0 && centroid) {
      const marker = L.circleMarker(centroid, {
        radius: 7, color: color, fillColor: color, fillOpacity: 0.5, weight: 2.5
      });
      marker.bindPopup(popupContent(item), { maxWidth: 320 });
      marker.addTo(layerGroup);
      polygonLayers.push({ layer: marker, item: item, idx: idx, centroid: centroid, isMarker: true });
    }

    // Lista lateral
    const div = document.createElement('div');
    div.className = 'list-item';

    const linked      = isLinked(item);
    const displayName = linked ? (item.e.empreendimento || item.e.nome || item.n) : item.n;
    const cityText    = linked ? (item.e.cidade || '') : '';
    const regional    = linked ? (item.e.regional || '') : '';
    const unitsText   = linked && item.e.total_unidades ? fmtNum(item.e.total_unidades) + ' un.' : '';
    const hasLocation = !!(centroid);

    div.innerHTML = `
      <div class="meta">
        ${regional ? `<span class="regional-tag" style="background:${colors[regional] || '#5a6e8e'}">${regional}</span>` : ''}
        ${cityText  ? `<span>${cityText}</span>`   : ''}
        ${unitsText ? `<span>${unitsText}</span>`  : ''}
        ${!linked       ? '<span class="no-match">sem dados</span>' : ''}
        ${!hasLocation  ? '<span class="no-match" title="Sem coordenadas cadastradas">sem loc.</span>' : ''}
      </div>
      <div class="name" title="${item.n}">${displayName}</div>`;

    div.onclick = () => {
      if (centroid) {
        map.flyTo(centroid, 14, { duration: 1.0 });
        const pl = polygonLayers.find(p => p.idx === idx);
        if (pl) setTimeout(() => pl.layer.openPopup(centroid), 600);
      }
      document.querySelectorAll('.list-item').forEach(el => el.classList.remove('highlight'));
      div.classList.add('highlight');
      if (window.innerWidth <= 768) closeSidebar();
    };

    listEl.appendChild(div);
  });

  // document.getElementById('counter').textContent = visibleCount + ' de ' + items.length + ' terrenos';
  document.getElementById('counter').textContent = `Lista de empreendimentos:`;

  buildOverviewMarkers();
  setTimeout(applyZoomVisibility, 0);
}

// Initial render
updateMap();

// Fix Leaflet map size on resize
window.addEventListener('resize', () => {
  setTimeout(() => map.invalidateSize(), 100);
});