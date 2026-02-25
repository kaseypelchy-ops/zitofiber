/* Zito FieldOS — application logic
 * Split from single-file build on 2026.02.24
 * To update: bump ?v= query string in index.html <script> tag
 */

// ──────────────────────────────────────────────────────────
//  STATE
// ──────────────────────────────────────────────────────────
var APP_NAME    = 'Zito FieldOS';
var APP_TAGLINE = 'Field Operations & Sales Intelligence';
var APP_VERSION = '1.0.3';
var BUILD_ID    = '2026.02.22';
var APP_ENV     = 'Production';

var addresses  = [];
var activeId   = null;
var selPkg     = null;
var selStatus  = null;
var selSlot    = null;
var webhookURL = 'https://script.google.com/macros/s/AKfycbyyqHh3H5qbBxB2fP9dPsymDoreXGwvrjCLT-ROQGBLMjBXKpprt3LWCC2aHbbeovJp/exec';
var repName    = 'Rep';
var repPhone   = '';
var repEmail   = '';
var repWebsite = 'https://www.zitomedia.net';
var activeTerritory = '';
var mapObj     = null;
var mapMarkers = {};
var kmlGeoJSON = null;
var toastTimer = null;
var sidebarOpen  = true;
var pinDropMode  = false;
var tempPinMarker = null;

// ──────────────────────────────────────────────────────────
//  COLORS
// ──────────────────────────────────────────────────────────
var COLORS = {
  pending:       '#6b7280',
  mega:          '#8b5cf6',
  gig:           '#10b981',
  nothome:       '#d97706',
  brightspeed:   '#ef4444',
  incontract:    '#818cf8',
  notinterested: '#dc2626',
  goback:        '#06b6d4',
  vacant:        '#ca8a04',
  business:      '#6366f1'
};
var COLOR_ACTIVE = '#facc15';
var COLOR_PASSED = '#6b7280';

var colors = {
  accent: '#005696',
  mega:   '#8b5cf6',
  gig:    '#10b981',
  warn:   '#d97706',
  danger: '#ef4444',
  muted:  '#8b949e'
};


// ──────────────────────────────────────────────────────────
//  SIDEBAR TOGGLE
// ──────────────────────────────────────────────────────────
function toggleSidebar() {
  sidebarOpen = !sidebarOpen;
  var app = document.getElementById('page-app');
  var btn = document.getElementById('sidebar-toggle');
  if (sidebarOpen) {
    app.classList.remove('sidebar-collapsed');
    btn.innerHTML = '&#8249;';
    btn.title = 'Hide address list';
  } else {
    app.classList.add('sidebar-collapsed');
    btn.innerHTML = '&#8250;';
    btn.title = 'Show address list';
  }
  if (mapObj) {
    setTimeout(function() { mapObj.invalidateSize(); }, 260);
  }
}

function maybeAutoCollapse() {
  if (window.innerWidth <= 640) {
    sidebarOpen = true;
    toggleSidebar();
  }
}

// ──────────────────────────────────────────────────────────
//  MODAL
// ──────────────────────────────────────────────────────────
function openModal()  { document.getElementById('modal').classList.add('open'); }
function closeModal() { document.getElementById('modal').classList.remove('open'); }
function handleModalClick(e) { if (e.target === document.getElementById('modal')) closeModal(); }

// ──────────────────────────────────────────────────────────
//  FILE INPUTS
// ──────────────────────────────────────────────────────────

// Lazy-load a script only when it's first needed.
// PapaParse (14KB) and JSZip (25KB) are skipped entirely on
// normal sessions where no file upload happens.
function lazyLoad(url, cb) {
  if (document.querySelector('script[src="' + url + '"]')) { cb(); return; }
  var s = document.createElement('script');
  s.src = url;
  s.onload = cb;
  document.head.appendChild(s);
}
var PAPAPARSE_URL = 'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js';
var JSZIP_URL     = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';

document.getElementById('csv-file').addEventListener('change', function() {
  var f = this.files[0];
  if (!f) return;
  var self = this;
  lazyLoad(PAPAPARSE_URL, function() {
  Papa.parse(f, {
    header: true,
    skipEmptyLines: true,
    complete: function(res) {
      addresses = [];
      res.data.forEach(function(row, i) {
        var keys = Object.keys(row);
        function col(names) {
          for (var n of names) {
            var k = keys.find(function(k){ return k.toLowerCase().trim() === n; });
            if (k !== undefined && row[k] !== undefined && String(row[k]).trim()) return String(row[k]).trim();
          }
          return '';
        }
        var addr = col(['address','street address','street']);
        if (!addr) return;
        var activeCount = col(['active count','active_count','activecount','active','type','customer type','customertype']).toLowerCase().trim();
        addresses.push({
          id: i,
          address: addr,
          city:  col(['city']),
          state: col(['state']),
          zip:   col(['zip','zipcode','zip code','postal','postal code']),
          lat:   parseFloat(col(['lat','latitude']))  || null,
          lng:   parseFloat(col(['lng','lon','longitude'])) || null,
          activeCount: activeCount,
          status: 'pending',
          sale: null
        });
      });
      var el = document.getElementById('csv-status');
      if (addresses.length > 0) {
        el.className = 'dz-status ok';
        el.textContent = '✓ ' + addresses.length + ' addresses loaded';
        checkLaunchReady();
      } else {
        el.className = 'dz-status err';
        el.textContent = '✗ No addresses found — check column names (need: address, city, state, zip)';
      }
    }
  });
  }); // end lazyLoad
});

// ── KMZ / KML ────────────────────────────────────────────
var kmlFiles = [];

document.getElementById('kml-file').addEventListener('change', function() {
  var files = Array.from(this.files);
  if (!files.length) return;
  var input = this;
  lazyLoad(JSZIP_URL, function() {
    files.forEach(function(f) { loadKmlFile(f); });
    input.value = '';
  });
});

function loadKmlFile(f) {
  var ext = f.name.split('.').pop().toLowerCase();
  if (ext === 'kmz') {
    JSZip.loadAsync(f).then(function(zip) {
      var kmlEntry = null;
      zip.forEach(function(path, file) {
        if (!kmlEntry && path.toLowerCase().endsWith('.kml')) kmlEntry = file;
      });
      if (!kmlEntry) { addKmlFileRow(f.name, [], '⚠ No KML inside'); return; }
      kmlEntry.async('string').then(function(text) {
        var features = parseKmlFeatures(text);
        addKmlFileRow(f.name, features, features.length ? null : '⚠ No polygons found');
      });
    }).catch(function() { addKmlFileRow(f.name, [], '⚠ Could not unzip'); });
  } else {
    var reader = new FileReader();
    reader.onload = function(e) {
      var features = parseKmlFeatures(e.target.result);
      addKmlFileRow(f.name, features, features.length ? null : '⚠ No polygons found');
    };
    reader.readAsText(f);
  }
}

function parseKmlFeatures(text) {
  try {
    var xml  = new DOMParser().parseFromString(text, 'text/xml');
    var feats = [];
    xml.querySelectorAll('coordinates').forEach(function(node) {
      var pts = node.textContent.trim().split(/\s+/).map(function(s) {
        var p = s.split(',');
        return [parseFloat(p[0]), parseFloat(p[1])];
      }).filter(function(p){ return !isNaN(p[0]) && !isNaN(p[1]); });
      if (pts.length > 2) {
        feats.push({ type:'Feature', geometry:{ type:'Polygon', coordinates:[pts] }, properties:{} });
      }
    });
    return feats;
  } catch(e) { return []; }
}

function addKmlFileRow(name, features, errMsg) {
  var ok = !errMsg && features.length > 0;
  var uid = 'kf-' + Date.now() + '-' + Math.random().toString(36).slice(2,6);
  if (ok) {
    kmlFiles.push({ uid: uid, name: name, features: features });
    rebuildKmlGeoJSON();
  }
  var list = document.getElementById('kml-file-list');
  var row  = document.createElement('div');
  row.className = 'kml-file-row ' + (ok ? 'ok' : 'err');
  row.id = uid;
  row.innerHTML =
    '<span class="kml-file-icon">' + (ok ? '🗺️' : '⚠️') + '</span>' +
    '<span class="kml-file-name" title="' + escHtml(name) + '">' + escHtml(name) + '</span>' +
    '<span class="kml-file-status">' + (ok ? features.length + ' polygon' + (features.length !== 1 ? 's' : '') : errMsg) + '</span>' +
    (ok ? '<button class="kml-file-remove" onclick="removeKmlFile(\'' + uid + '\')" title="Remove">✕</button>' : '');
  list.appendChild(row);
}

function removeKmlFile(uid) {
  kmlFiles = kmlFiles.filter(function(f){ return f.uid !== uid; });
  var row = document.getElementById(uid);
  if (row) row.remove();
  rebuildKmlGeoJSON();
}

function rebuildKmlGeoJSON() {
  var allFeatures = [];
  kmlFiles.forEach(function(f){ allFeatures = allFeatures.concat(f.features); });
  kmlGeoJSON = allFeatures.length > 0
    ? { type:'FeatureCollection', features: allFeatures }
    : null;
}

['dz-csv','dz-kml'].forEach(function(id) {
  var el = document.getElementById(id);
  el.addEventListener('dragover',  function(e){ e.preventDefault(); el.classList.add('dz-over'); });
  el.addEventListener('dragleave', function(){ el.classList.remove('dz-over'); });
  el.addEventListener('drop',      function(){ el.classList.remove('dz-over'); });
});

// ──────────────────────────────────────────────────────────
//  LOAD ADDRESSES FROM SHEET
// ──────────────────────────────────────────────────────────
function fetchAddressesFromSheet() {
  var btn = document.getElementById('btn-fetch-addr');
  var st  = document.getElementById('fetch-addr-status');
  var repInput = (document.getElementById('rep-name') ? (document.getElementById('rep-name').value || '').trim() : '');
  if (!repInput || repInput.split(/\s+/).filter(function(p){ return p.length > 0; }).length < 2) {
    st.className = 'dz-status err';
    st.textContent = '✗ Enter your full name first (First Last).';
    return;
  }

  btn.disabled = true;
  document.getElementById('fetch-addr-icon').textContent = '⏳';
  st.className = 'dz-status';
  st.textContent = 'Loading addresses…';

  // Let the backend know this is a manager so it returns ALL territories
  var managerFlag = MANAGER_NAMES.indexOf(repInput.toLowerCase()) >= 0 ? '&isManager=true' : '';
  fetch(webhookURL + '?action=addresses&repName=' + encodeURIComponent(repInput) + managerFlag + '&_t=' + Date.now())
    .then(function(r){ return r.json(); })
    .then(function(json){
      if (!json || !json.rows) throw new Error('Bad response from server');
      if (json.status === 'error') throw new Error(json.message || 'Server error');

  activeTerritory = (json.territory || '').trim();

addresses = json.rows.map(function(row, i) {
  var lat = (row.lat !== '' && row.lat != null) ? parseFloat(row.lat) : null;
  var lng = (row.lng !== '' && row.lng != null) ? parseFloat(row.lng) : null;

  return {
    id:           i,
    sheetRow:     row.sheetRow,
    territory:    (row.territory || activeTerritory || '').trim(),
    address:      (row.address || '').trim(),
    city:         (row.city || '').trim(),
    state:        (row.state || '').trim(),
    zip:          (row.zip || '').trim(),

    lat:          (isFinite(lat) ? lat : null),
    lng:          (isFinite(lng) ? lng : null),

    activeCount:  (row.activeCount || row.active_count || row.type || '').toString().trim(),
    status:       (row.status || 'pending').toLowerCase(),
    salesperson:  (row.salesperson || '').trim(),

    // ✅ IMPORTANT: bring note over from Apps Script
    note:         (row.note || row.dispositionNote || row.disposition_note || '').toString().trim(),

    sale:         null
  };
});

// ✅ These are the missing pieces that make sidebar + map update immediately
updateStats();
buildList();
refreshMapMarkers(); // if you have this helper; otherwise use the fallback below

// Fallback if you DO NOT have refreshMapMarkers():
// if (mapObj) {
//   // clear old markers if you track them
//   if (mapMarkers) {
//     Object.keys(mapMarkers).forEach(function(k){
//       try { mapObj.removeLayer(mapMarkers[k]); } catch(e) {}
//     });
//     mapMarkers = {};
//   }
//   addresses.forEach(function(a){ if (a.lat != null && a.lng != null) placeMarker(a); });
// }

st.className   = 'dz-status ok';
st.textContent = '✓ ' + addresses.length + ' addresses loaded' + (activeTerritory ? (' • ' + activeTerritory) : '');
document.getElementById('fetch-addr-icon').textContent = '✅';
btn.disabled = false;
checkLaunchReady();

      st.className   = 'dz-status ok';
      st.textContent = '✓ ' + addresses.length + ' addresses loaded' + (activeTerritory ? (' • ' + activeTerritory) : '');
      document.getElementById('fetch-addr-icon').textContent = '✅';
      btn.disabled = false;
      checkLaunchReady();
    })
    .catch(function(err){
      st.className   = 'dz-status err';
      st.textContent = '✗ ' + (err && err.message ? err.message : 'Unable to load addresses');
      document.getElementById('fetch-addr-icon').textContent = '📋';
      btn.disabled = false;
    });
}


// ──────────────────────────────────────────────────────────
//  NAME VALIDATION
// ──────────────────────────────────────────────────────────
function hasValidName() {
  var val   = (document.getElementById('rep-name').value || '').trim();
  var parts = val.split(/\s+/).filter(function(p){ return p.length > 0; });
  return parts.length >= 2 && val.toLowerCase() !== 'rep';
}

function validateRepName() {
  var hint = document.getElementById('rep-name-hint');
  var val  = (document.getElementById('rep-name').value || '').trim();
  if (val.length > 0 && !hasValidName()) {
    hint.style.display = 'block';
  } else {
    hint.style.display = 'none';
  }
  checkLaunchReady();
}

function checkLaunchReady() {
  var hasAddresses = addresses.length > 0;
  document.getElementById('launch-btn').disabled = !(hasAddresses && hasValidName());
}

// ──────────────────────────────────────────────────────────
//  REAL-TIME POLLING
// ──────────────────────────────────────────────────────────
var pollTimer = null;

function startPolling() {
  // Managers always poll — they have no single activeTerritory
  if (!activeTerritory && !isManager()) return;

  pollTimer = setInterval(function() {
    var pollUrl = isManager()
      ? webhookURL + '?action=addresses&isManager=true&_t=' + Date.now()
      : webhookURL + '?action=addresses&territory=' + encodeURIComponent(activeTerritory || '') + '&_t=' + Date.now();
    fetch(pollUrl)
      .then(function(r){ return r.json(); })
      .then(function(json){
        if (!json || !json.rows) return;

        var changed = false;

        json.rows.forEach(function(row) {
          var addr = addresses.find(function(a){ return a.sheetRow === row.sheetRow; });
          if (!addr) return;

          // Status update
          var newStatus = (row.status || 'pending').toString().toLowerCase().trim();
          if (addr.status !== newStatus) {
            addr.status = newStatus;
            changed = true;
          }

          // Note update (independent of status)
          var newNote = (row.note || row.dispositionNote || row.disposition_note || '').toString().trim();
          if (addr.note !== newNote) {
            addr.note = newNote;
            changed = true;
          }

          // If anything changed, refresh this marker
          if (changed && addr.lat && addr.lng) {
            placeMarker(addr);
          }
        });

        if (changed) {
          buildList();
          updateStats();
        }
      })
      .catch(function(){});
  }, 30000);
}
function launchApp() {
  repName = (document.getElementById('rep-name').value || '').trim();
  repPhone = (document.getElementById('rep-phone') ? (document.getElementById('rep-phone').value || '').trim() : '');
  repEmail = (document.getElementById('rep-email') ? (document.getElementById('rep-email').value || '').trim() : '');

  try {
    localStorage.setItem('zito_rep_name', repName);
    localStorage.setItem('zito_rep_phone', repPhone);
    localStorage.setItem('zito_rep_email', repEmail);
    if (!localStorage.getItem('fieldos_session_start')) {
      localStorage.setItem('fieldos_session_start', new Date().toISOString());
    }
  } catch(e) {}

  var splash = document.getElementById('splash');
  document.getElementById('splash-rep-name').textContent = repName;
  var fill = document.getElementById('splash-prog-fill');
  if (fill) { fill.style.animation = 'none'; fill.offsetHeight; fill.style.animation = ''; }
  if (splash) splash.classList.remove('gone', 'fade-out');

  var fadeTimer = setTimeout(function() {
    if (!splash) return;
    splash.classList.add('fade-out');
    setTimeout(function() { splash.classList.add('gone'); }, 700);
  }, 4500);

  try {
    document.getElementById('page-setup').style.display = 'none';
    document.getElementById('page-app').style.display   = 'block';

    updateStats();
    buildList();
    initMap();
    geocodeAll();
    startPolling();
    maybeAutoCollapse();
    initBadge();
    // Managers land on the team dashboard automatically
    if (isManager()) {
      setTimeout(function(){ openManagerPanel(); }, 600);
    }
  } catch (err) {
    try { clearTimeout(fadeTimer); } catch(e) {}
    if (splash) { splash.classList.add('fade-out'); setTimeout(function(){ splash.classList.add('gone'); }, 300); }
    console.error(err);
    toast('App error: ' + String(err), 't-err');
  }
}

// ──────────────────────────────────────────────────────────
//  MAP
// ──────────────────────────────────────────────────────────
  var wxRadarLayer = null;
  var wxRadarMeta  = null;
  var wxRadarOn    = false;
  var wxRadarRefreshTimer = null;

  var wxLastTempFetch = 0;
  var wxTempTimer = null;

  function wxSetRadarUI_(on) {
    wxRadarOn = !!on;
    var btn = document.getElementById('wx-radar-toggle');
    if (!btn) return;
    btn.classList.toggle('on', wxRadarOn);
    btn.setAttribute('aria-pressed', wxRadarOn ? 'true' : 'false');
  }

  function wxToggleRadar() {
    if (!mapObj) return;
    if (!wxRadarLayer) {
      wxSetRadarUI_(true);
      wxInitRadarOverlay_();
      return;
    }
    if (mapObj.hasLayer(wxRadarLayer)) {
      mapObj.removeLayer(wxRadarLayer);
      wxSetRadarUI_(false);
    } else {
      wxRadarLayer.addTo(mapObj);
      wxSetRadarUI_(true);
    }
  }

  function wxInitRadarOverlay_() {
    if (!mapObj) return;

    fetch('https://api.rainviewer.com/public/weather-maps.json')
      .then(function(r){ return r.json(); })
      .then(function(meta){
        wxRadarMeta = meta;
        var past = meta && meta.radar && meta.radar.past ? meta.radar.past : [];
        if (!past.length) return;

        var frame = past[past.length - 1];
        var host  = meta.host;
        var path  = frame.path;

        var tileUrl = host + path + '/256/{z}/{x}/{y}/2/1_1.png';

        var wasOn = wxRadarLayer && mapObj.hasLayer(wxRadarLayer);

        if (wxRadarLayer) {
          try { mapObj.removeLayer(wxRadarLayer); } catch(e) {}
        }

        wxRadarLayer = L.tileLayer(tileUrl, {
          opacity: 0.55,
          zIndex: 500,
          maxNativeZoom: 7,
          maxZoom: 19
        });

        if (wasOn || wxRadarOn) {
          wxRadarLayer.addTo(mapObj);
          wxSetRadarUI_(true);
        }
      })
      .catch(function(){});
  }

  function wxFetchTemp_(lat, lng) {
    var now = Date.now();
    if (now - wxLastTempFetch < 60 * 1000) return;
    wxLastTempFetch = now;

    var url =
      'https://api.open-meteo.com/v1/forecast' +
      '?latitude=' + encodeURIComponent(lat) +
      '&longitude=' + encodeURIComponent(lng) +
      '&current_weather=true' +
      '&temperature_unit=fahrenheit';

    fetch(url)
      .then(function(r){ return r.json(); })
      .then(function(json){
        var el = document.getElementById('wx-temp');
        if (!el) return;

        var t = (json && json.current_weather && typeof json.current_weather.temperature === 'number')
          ? json.current_weather.temperature
          : null;

        if (t === null) { el.textContent = '—°F'; return; }
        el.textContent = Math.round(t) + '°F';
      })
      .catch(function(){
        var el = document.getElementById('wx-temp');
        if (el) el.textContent = '—°F';
      });
  }

  function wxUpdateTempFromMap_() {
    if (!mapObj) return;
    var c = mapObj.getCenter();
    wxFetchTemp_(c.lat, c.lng);
  }

  function wxInitTemperature_() {
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(function(pos){
        wxFetchTemp_(pos.coords.latitude, pos.coords.longitude);
      }, function(){
        wxUpdateTempFromMap_();
      }, { enableHighAccuracy: false, timeout: 5000, maximumAge: 600000 });
    } else {
      wxUpdateTempFromMap_();
    }

    var t;
    mapObj.on('moveend', function(){
      clearTimeout(t);
      t = setTimeout(wxUpdateTempFromMap_, 700);
    });

    if (wxTempTimer) clearInterval(wxTempTimer);
    wxTempTimer = setInterval(wxUpdateTempFromMap_, 10 * 60 * 1000);
  }

// Track the active base layer and label overlay globally
var activeBaseLayer   = null;
var activeLabelLayer  = null;

// Satellite imagery (ONLY base map option) + labels overlay
var SATELLITE_LAYER = {
  url:  'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
  opts: {
    attribution: '© Esri, Maxar, Earthstar Geographics, and the GIS User Community',
    maxZoom: 20,
    maxNativeZoom: 19
  }
};

// Reference labels so streets/places are readable on imagery
var LABELS_LAYER = {
  url:  'https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}',
  opts: {
    attribution: '',
    maxZoom: 20,
    maxNativeZoom: 19,
    pane: 'overlayPane'
  }
};

function setSatelliteBaseLayer() {
  if (!mapObj) return;

  if (activeBaseLayer)  { mapObj.removeLayer(activeBaseLayer);  activeBaseLayer = null; }
  if (activeLabelLayer) { mapObj.removeLayer(activeLabelLayer); activeLabelLayer = null; }

  activeBaseLayer  = L.tileLayer(SATELLITE_LAYER.url, SATELLITE_LAYER.opts).addTo(mapObj);
  activeLabelLayer = L.tileLayer(LABELS_LAYER.url, LABELS_LAYER.opts).addTo(mapObj);
}

function initMap() {
  mapObj = L.map('map');

  // Default to satellite — best for pin dropping on houses
  // Base map (Satellite imagery only)
  setSatelliteBaseLayer();

  mapObj.setView([39.5, -98.35], 5);

  if (kmlGeoJSON && kmlGeoJSON.features.length > 0) {
    var palette = [
      { stroke:'#2563eb', fill:'#3b82f6' },
      { stroke:'#d97706', fill:'#f59e0b' },
      { stroke:'#059669', fill:'#10b981' },
      { stroke:'#dc2626', fill:'#ef4444' },
      { stroke:'#7c3aed', fill:'#8b5cf6' },
      { stroke:'#0891b2', fill:'#06b6d4' },
    ];
    var allBounds = [];
    kmlFiles.forEach(function(kf, i) {
      if (!kf.features.length) return;
      var col = palette[i % palette.length];
      var layer = L.geoJSON({ type:'FeatureCollection', features: kf.features }, {
        style: { color: col.stroke, weight: 3, fillColor: col.fill, fillOpacity: 0.12, dashArray: '8 4' }
      }).addTo(mapObj);
      allBounds.push(layer.getBounds());
    });
    if (allBounds.length) {
      var combined = allBounds[0];
      allBounds.forEach(function(b){ combined.extend(b); });
      setTimeout(function(){ mapObj.fitBounds(combined, { padding:[40,40] }); }, 100);
    }
  }

  addresses.forEach(function(a) {
    if (a.lat && a.lng) { placeMarker(a); }
  });

  // Fit map to address pins if we have any, otherwise fall back to US overview.
  // KML bounds take priority if territories were loaded.
  if (!kmlGeoJSON || !kmlGeoJSON.features.length) {
    fitToAddresses();
  }

  wxSetRadarUI_(false);
  wxInitRadarOverlay_();
  if (wxRadarRefreshTimer) clearInterval(wxRadarRefreshTimer);
  wxRadarRefreshTimer = setInterval(wxInitRadarOverlay_, 10 * 60 * 1000);

  wxInitTemperature_();

  // ── Pin-drop: tap map to place a new address ──────────────
  mapObj.on('click', function(e) {
    if (!pinDropMode) return;
    handleMapPinDrop(e.latlng);
  });

  // ── Leaflet custom control: Drop Pin button ───────────────
  var PinDropControl = L.Control.extend({
    options: { position: 'bottomleft' },
    onAdd: function() {
      var container = L.DomUtil.create('div', 'leaflet-bar leaflet-control pin-drop-control');
      var btn = L.DomUtil.create('a', 'pin-drop-btn', container);
      btn.id          = 'btn-pin-drop';
      btn.href        = '#';
      btn.title       = 'Drop a pin to add a home';
      btn.innerHTML   = '<span class="pin-drop-icon">📍</span><span class="pin-drop-label">Drop Pin</span>';
      btn.setAttribute('role', 'button');
      btn.setAttribute('aria-label', 'Drop pin to add address');

      L.DomEvent.on(btn, 'click', function(e) {
        L.DomEvent.stopPropagation(e);
        L.DomEvent.preventDefault(e);
        togglePinDropMode();
      });
      // Prevent map drag from starting on this control
      L.DomEvent.disableClickPropagation(container);
      L.DomEvent.disableScrollPropagation(container);
      return container;
    }
  });
  new PinDropControl().addTo(mapObj);

  // ── Map style switcher control ────────────────────────────
    // (Map switcher removed: Voyager is the only base map)

}

function getMarkerColor(addr) {
  var s = (addr.status || '').toLowerCase().trim();
  if (COLORS[s]) return COLORS[s];
  var shape = getMarkerShape(addr);
  if (shape === 'bolt')  return COLOR_ACTIVE;
  if (shape === 'house') return COLOR_PASSED;
  return COLORS.pending;
}

function markerHTML(color, shape) {
  if (shape === 'house') {
    return '<div style="width:26px;height:26px;background:' + color + ';clip-path:polygon(50% 0%,100% 45%,85% 45%,85% 100%,15% 100%,15% 45%,0% 45%);filter:drop-shadow(0 2px 3px rgba(0,0,0,0.55))"></div>';
  }
  if (shape === 'bolt') {
    return '<div style="width:20px;height:28px;background:' + color + ';clip-path:polygon(65% 0%,20% 52%,48% 52%,35% 100%,80% 42%,52% 42%,68% 0%);filter:drop-shadow(0 2px 3px rgba(0,0,0,0.55))"></div>';
  }
  return '<div style="width:16px;height:16px;border-radius:50%;background:' + color + ';border:2.5px solid #fff;box-shadow:0 2px 5px rgba(0,0,0,0.5)"></div>';
}

// ── FIX: Rep-logged no-sale statuses are always 'dot', never 'bolt' ──────────
function getMarkerShape(addr) {
  var s  = (addr.status      || '').toLowerCase().trim();
  var ac = (addr.activeCount || '').toLowerCase().trim();

  // Sales outcomes always use explicit shapes regardless of activeCount
  if (s === 'mega' || s === 'gig') return 'house';

  // Rep-logged no-sale statuses: always a dot — NEVER treat as active customer.
  // Without this guard, any address with a non-empty activeCount field would
  // fall through to the `if (ac && ac !== '') return 'bolt'` catch-all below,
  // incorrectly showing "Active Customer" after Go Back Later / Not Interested /
  // Brightspeed etc. are submitted.
  var REP_LOGGED = ['nothome','brightspeed','incontract','notinterested','goback','vacant','business'];
  if (REP_LOGGED.indexOf(s) >= 0) return 'dot';

  // Sheet-driven status / activeCount checks (untouched addresses only)
  if (s === 'active') return 'bolt';
  if (s.indexOf('home') >= 0 || s.indexOf('passed') >= 0) return 'house';
  if (ac === 'active' || ac === 'existing' || ac === 'customer') return 'bolt';
  if (ac.indexOf('home') >= 0 || ac.indexOf('passed') >= 0 || ac === 'hp') return 'house';
  if (ac && ac !== '') return 'bolt';
  return 'dot';
}

function placeMarker(addr) {
  if (mapMarkers[addr.id]) { mapMarkers[addr.id].remove(); delete mapMarkers[addr.id]; }

  var color = getMarkerColor(addr);
  var shape = getMarkerShape(addr);
  var html  = markerHTML(color, shape);
  var size   = shape === 'house' ? [26,26] : shape === 'bolt' ? [20,28] : [16,16];
  var anchor = shape === 'house' ? [13,26] : shape === 'bolt' ? [10,28] : [8,8];
  var icon  = L.divIcon({ className:'', html: html, iconSize: size, iconAnchor: anchor });
  var m     = L.marker([addr.lat, addr.lng], { icon: icon }).addTo(mapObj);

  var sub = [addr.city, addr.state, addr.zip].filter(Boolean).join(', ');
  var typeLabel = shape === 'bolt' ? '⚡ Active Customer' : (shape === 'house' ? '🏠 Homes Passed' : '');
  var pid = addr.id;

  var btnHTML = shape === 'bolt'
    ? '<button class="pop-open-btn pop-active-btn" onclick="openFormFromMap(' + pid + ')">⚡ View Address</button>'
    : '<button class="pop-open-btn" onclick="openFormFromMap(' + pid + ')">Open Sales Form</button>';

  // ✅ NEW
  var noteHTML = (addr.note && addr.note.trim())
    ? '<div style="margin-top:6px;padding:6px 8px;border:1px solid #e5e7eb;border-radius:10px;background:#f9fafb;font-size:11px;color:#111;"><b>Note:</b> ' + escHtml(addr.note) + '</div>'
    : '';

m.bindPopup(
  '<div style="font-family:Syne,sans-serif;min-width:160px">' +
    popupHtmlForAddr(addr) +
    btnHTML +
  '</div>',
  { minWidth: 180 }
);

  mapMarkers[addr.id] = m;
}

window.openFormFromMap = function(id) {
  if (mapObj) mapObj.closePopup();
  openForm(id);
};
// ──────────────────────────────────────────────────────────
//  GEOCODING
// ──────────────────────────────────────────────────────────
function fitToAddresses() {
  if (!mapObj) return;
  var pinned = addresses.filter(function(a) { return a.lat && a.lng; });
  if (pinned.length === 0) {
    mapObj.setView([39.5, -98.35], 5); // no pins yet — show whole US
    return;
  }
  if (pinned.length === 1) {
    // Single pin — go straight to street level
    mapObj.setView([pinned[0].lat, pinned[0].lng], 17);
    return;
  }
  // Multiple pins — fit all of them with padding, then cap zoom at 17
  // so we don't land on a comically close view when all pins are on one street
  var bounds = L.latLngBounds(pinned.map(function(a) { return [a.lat, a.lng]; }));
  mapObj.fitBounds(bounds, { padding: [48, 48], maxZoom: 17 });
}

function geocodeAll() {
  var toGeocode = addresses.filter(function(a) { return !a.lat || !a.lng; });
  if (toGeocode.length === 0) { buildList(); return; }

  var total   = toGeocode.length;
  var done    = 0;
  var failed  = 0;
  showGeocodeBar(done, total);

  var idx = 0;

  function geocodeOne(a) {
    var query = [a.address, a.city, a.state, a.zip].filter(Boolean).join(', ');
    var url   = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=us&q=' + encodeURIComponent(query);

    fetch(url, { headers: { 'Accept': 'application/json', 'User-Agent': 'FieldSalesApp/1.0' } })
      .then(function(r) { return r.json(); })
      .then(function(data) {
        if (data && data.length > 0) {
          a.lat = parseFloat(data[0].lat);
          a.lng = parseFloat(data[0].lon);
          if (mapObj) placeMarker(a);
        } else {
          failed++;
          a._geocodeFailed = true;
        }
        done++;
        showGeocodeBar(done, total, failed);
        buildList();
        scheduleNext();
      })
      .catch(function() {
        failed++;
        a._geocodeFailed = true;
        done++;
        showGeocodeBar(done, total, failed);
        scheduleNext();
      });
  }

  function scheduleNext() {
    if (idx < toGeocode.length) {
      var a = toGeocode[idx++];
      setTimeout(function() { geocodeOne(a); }, 1100);
    } else if (done >= total) {
      if (failed > 0) {
        document.getElementById('gc-text').textContent =
          '⚠ ' + (total - failed) + '/' + total + ' geocoded. ' + failed + ' not found.';
        document.getElementById('gc-fill').style.background = '#d97706';
        setTimeout(hideGeocodeBar, 6000);
      } else {
        document.getElementById('gc-text').textContent = '✓ All ' + total + ' addresses geocoded';
        document.getElementById('gc-fill').style.background = '#059669';
        setTimeout(hideGeocodeBar, 2500);
        fitToAddresses();
      }
    }
  }

  scheduleNext();
  setTimeout(scheduleNext, 1100);
}

function showGeocodeBar(done, total, failed) {
  var bar = document.getElementById('geocode-bar');
  if (!bar) return;
  failed = failed || 0;
  var pct = Math.round((done / total) * 100);
  bar.style.display = 'flex';
  var found = done - failed;
  document.getElementById('gc-text').textContent = 'Geocoding… ' + found + ' found, ' + failed + ' not found — ' + done + '/' + total;
  document.getElementById('gc-fill').style.width = pct + '%';
}

function hideGeocodeBar() {
  var bar = document.getElementById('geocode-bar');
  if (bar) bar.style.display = 'none';
}

// ──────────────────────────────────────────────────────────
//  ADDRESS LIST
// ──────────────────────────────────────────────────────────
var TAG_HTML  = {
  mega:          '<span class="ar-tag tag-mega">⚡ Mega</span>',
  gig:           '<span class="ar-tag tag-gig">🚀 Gig</span>',
  nothome:       '<span class="ar-tag tag-nh">🚪 Not Home</span>',
  brightspeed:   '<span class="ar-tag tag-bs">⚡ Brightspeed</span>',
  incontract:    '<span class="ar-tag tag-ic">📋 In Contract</span>',
  notinterested: '<span class="ar-tag tag-ni">❌ Not Int.</span>',
  goback:        '<span class="ar-tag tag-gbl">🔄 Go Back</span>',
  vacant:        '<span class="ar-tag tag-vac">🏚️ Vacant</span>',
  business:      '<span class="ar-tag tag-biz">🏢 Business</span>'
};

function buildList(filter) {
  var list = filter
    ? addresses.filter(function(a) {
        var q = filter.toLowerCase();
        return a.address.toLowerCase().indexOf(q) >= 0 ||
               (a.city && a.city.toLowerCase().indexOf(q) >= 0) ||
               (a.zip  && a.zip.indexOf(q) >= 0);
      })
    : addresses;

  document.getElementById('addr-count').textContent = addresses.length;

  var html = list.map(function(a) {
    var sub   = [a.city, a.state, a.zip].filter(Boolean).join(', ') || '—';
    var tag   = TAG_HTML[a.status] || '';
    var selC  = (a.id === activeId) ? ' sel' : '';
    var color = getMarkerColor(a);
    var shape = getMarkerShape(a);
    var icon;
    if (shape === 'bolt') {
      icon = '<div style="width:11px;height:15px;background:' + color + ';clip-path:polygon(65% 0%,20% 52%,48% 52%,35% 100%,80% 42%,52% 42%,68% 0%)"></div>';
    } else if (shape === 'house') {
      icon = '<div style="width:14px;height:14px;background:' + color + ';clip-path:polygon(50% 0%,100% 45%,85% 45%,85% 100%,15% 100%,15% 45%,0% 45%)"></div>';
    } else {
      icon = '<div style="width:10px;height:10px;border-radius:50%;background:' + color + '"></div>';
    }
var noteLine = (a.note && a.note.trim())
  ? '<div class="ar-note">' + escHtml(a.note.trim()) + '</div>'
  : '';

return '<div class="addr-row' + selC + '" data-id="' + a.id + '">' +
  '<div class="ar-dot">' + icon + '</div>' +
  '<div class="ar-info">' +
    '<div class="ar-st">'  + escHtml(a.address) + '</div>' +
    '<div class="ar-sub">' + escHtml(sub)        + '</div>' +
    noteLine +
  '</div>' + tag + '</div>';
  }).join('');

  var container = document.getElementById('addr-items');
  container.innerHTML = html || '<div style="padding:24px;text-align:center;color:var(--muted);font-size:12px">No addresses found</div>';

  container.querySelectorAll('.addr-row').forEach(function(row) {
    row.addEventListener('click', function() {
      var id = parseInt(this.getAttribute('data-id'), 10);
      openForm(id);
      if (window.innerWidth <= 640 && sidebarOpen) toggleSidebar();
    });
  });
}

function filterList(val) { buildList(val || null); }

function escHtml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ──────────────────────────────────────────────────────────
//  FORM
// ──────────────────────────────────────────────────────────
function openForm(id) {
  var addr = null;
  for (var i = 0; i < addresses.length; i++) { if (addresses[i].id === id) { addr = addresses[i]; break; } }
  if (!addr) return;

  setFormCollapsed(false);

  if (getMarkerShape(addr) === 'bolt') {
    activeId = id;
    document.getElementById('pf-addr-line').textContent = addr.address;
    document.getElementById('pf-addr-sub').textContent  = [addr.city, addr.state, addr.zip].filter(Boolean).join(', ');
    document.getElementById('active-notice').style.display = 'block';
    document.getElementById('sales-form-body').style.display = 'none';
    document.getElementById('panel-form').classList.add('open');
    document.body.classList.add('form-open');
    buildList();
    return;
  }

  document.getElementById('active-notice').style.display = 'none';
  document.getElementById('sales-form-body').style.display = 'block';

  activeId = id;

  document.getElementById('pf-addr-line').textContent = addr.address;
  document.getElementById('pf-addr-sub').textContent  = [addr.city, addr.state, addr.zip].filter(Boolean).join(', ');

  var s = addr.sale || {};
  document.getElementById('f-first').value = s.firstName || '';
  document.getElementById('f-last').value  = s.lastName  || '';
  document.getElementById('f-phone').value = s.phone     || '';
  document.getElementById('f-email').value = s.email     || '';
  document.getElementById('f-notes').value = s.notes     || '';

  selPkg    = null;
  selStatus = null;
  document.getElementById('pkg-mega').className = 'pkg-card mega-card';
  document.getElementById('pkg-gig').className  = 'pkg-card gig-card';
  document.getElementById('btn-mega').disabled  = true;
  document.getElementById('btn-gig').disabled   = true;
  document.getElementById('btn-mega').textContent = '⚡ Submit — Mega Speed';
  document.getElementById('btn-gig').textContent  = '🚀 Submit — Gig Speed';
  document.getElementById('pricing-box').style.display        = 'none';
  document.getElementById('proration-section').style.display  = 'none';
  document.getElementById('sched-confirmed').style.display    = 'none';
  document.getElementById('sched-picker').style.display       = 'none';
  document.getElementById('sched-loading').style.display      = 'none';
  document.getElementById('sched-error').style.display        = 'none';
  document.getElementById('f-install-date').value = '';
  document.getElementById('f-install-time').value = '';
  selSlot = null;

  ['sbt-nh','sbt-bs','sbt-ic','sbt-ni','sbt-gbl','sbt-vac','sbt-biz'].forEach(function(sid) {
    document.getElementById(sid).className = 'stbtn';
  });

  // ── Restore previous disposition if address was already visited ──────────
  var statusToLabel = {
    nothome:       'Not Home',
    brightspeed:   'Brightspeed',
    incontract:    'In Contract',
    notinterested: 'Not Interested',
    goback:        'Go Back Later',
    vacant:        'Vacant',
    business:      'Business'
  };
  var statusToBtn = {
    nothome:       'sbt-nh',
    brightspeed:   'sbt-bs',
    incontract:    'sbt-ic',
    notinterested: 'sbt-ni',
    goback:        'sbt-gbl',
    vacant:        'sbt-vac',
    business:      'sbt-biz'
  };
  var statusCls = {
    nothome:       'act-nc',
    brightspeed:   'act-ni',
    incontract:    'act-vm',
    notinterested: 'act-ni',
    goback:        'act-cb',
    vacant:        'act-nc',
    business:      'act-vm'
  };
  var prevDisp    = document.getElementById('prev-disposition');
  var prevStatus  = document.getElementById('prev-disp-status');
  var prevNote    = document.getElementById('prev-disp-note');
  var nsWrap      = document.getElementById('ns-note-wrap');
  var nsNote      = document.getElementById('ns-note');
  var curStatus   = (addr.status || '').toLowerCase().trim();
  var curNote     = (addr.note   || '').trim();
  var prevLabel   = statusToLabel[curStatus];

  if (prevLabel) {
    // Show the banner
    prevStatus.textContent = prevLabel;
    prevStatus.className   = 'prev-disp-status s-' + curStatus;
    if (curNote) {
      prevNote.textContent   = '💬 ' + curNote;
      prevNote.style.display = 'block';
    } else {
      prevNote.style.display = 'none';
    }
    prevDisp.style.display = 'block';

    // Pre-highlight the matching status button
    selStatus = prevLabel;
    if (statusToBtn[curStatus]) {
      document.getElementById(statusToBtn[curStatus]).className = 'stbtn ' + statusCls[curStatus];
    }

    // Pre-fill note textarea (show it if this status normally has one)
    var needsNote = (prevLabel === 'Not Home' || prevLabel === 'Not Interested' || prevLabel === 'Go Back Later');
    if (nsWrap && nsNote) {
      nsWrap.style.display = needsNote ? 'block' : 'none';
      nsNote.value = curNote;
      if (needsNote) {
        nsNote.placeholder =
          (prevLabel === 'Not Home') ? 'Example: will return after 5pm / left flyer' :
          (prevLabel === 'Go Back Later') ? 'Example: customer asked to come back Friday' :
          'Example: not interested — already has provider';
      }
    }
  } else {
    // No prior no-sale disposition — hide banner, reset note
    prevDisp.style.display = 'none';
    if (nsWrap && nsNote) { nsWrap.style.display = 'none'; nsNote.value = ''; }
  }

  document.getElementById('panel-form').classList.add('open');
  document.body.classList.add('form-open');

  if (addr.lat && addr.lng && mapObj) {
    mapObj.panTo([addr.lat, addr.lng], { animate: true });
  }

  buildList();
}

function closeForm() {
  document.getElementById('panel-form').classList.remove('open');
  document.body.classList.remove('form-open');
  activeId  = null;
  selPkg    = null;
  selStatus = null;
  buildList();
}

function clearPrevDisposition() {
  var addr = getAddr();
  if (!addr) return;
  addr.status = 'pending';
  addr.note   = '';
  // Reset banner
  document.getElementById('prev-disposition').style.display = 'none';
  // Reset status buttons and note textarea
  ['sbt-nh','sbt-bs','sbt-ic','sbt-ni','sbt-gbl','sbt-vac','sbt-biz'].forEach(function(sid) {
    document.getElementById(sid).className = 'stbtn';
  });
  selStatus = null;
  var nsWrap = document.getElementById('ns-note-wrap');
  var nsNote = document.getElementById('ns-note');
  if (nsWrap) nsWrap.style.display = 'none';
  if (nsNote) nsNote.value = '';
  // Update marker and sidebar to reflect cleared status
  if (addr.lat && addr.lng) placeMarker(addr);
  buildList();
  updateAddressStatus(addr, 'pending', '');
  toast('🗑 Disposition cleared', 't-info');
}

// ──────────────────────────────────────────────────────────
//  SALES FORM COLLAPSE / EXPAND
// ──────────────────────────────────────────────────────────
var formCollapsed = false;

function setFormCollapsed(collapsed) {
  formCollapsed = !!collapsed;
  var body = document.querySelector('#panel-form .pf-body');
  var btn  = document.getElementById('pf-collapse-btn');
  if (!body || !btn) return;
  body.style.display = formCollapsed ? 'none' : 'block';
  btn.textContent = formCollapsed ? '▸' : '▾';
  btn.setAttribute('aria-expanded', String(!formCollapsed));
}

function toggleFormCollapse() {
  setFormCollapsed(!formCollapsed);
}

function pickPkg(p) {
  selPkg = p;
  document.getElementById('pkg-mega').className = 'pkg-card mega-card' + (p === 'mega' ? ' active' : '');
  document.getElementById('pkg-gig').className  = 'pkg-card gig-card'  + (p === 'gig'  ? ' active' : '');
  document.getElementById('btn-mega').disabled  = (p !== 'mega');
  document.getElementById('btn-gig').disabled   = (p !== 'gig');
  document.getElementById('pricing-box').style.display = 'block';
  schedShow();
  calcPricing();
}

// ──────────────────────────────────────────────────────────
//  SCHEDULE PICKER
// ──────────────────────────────────────────────────────────
var SCHED_URL    = 'https://script.google.com/macros/s/AKfycbyyqHh3H5qbBxB2fP9dPsymDoreXGwvrjCLT-ROQGBLMjBXKpprt3LWCC2aHbbeovJp/exec';
var SLOT_TIMES   = ['8:00 AM','10:00 AM','1:00 PM','3:00 PM'];
var schedData    = {};
var schedWeekOff = 0;

function schedNormalizeTime(raw) {
  if (!raw) return '';
  var s = String(raw).trim();
  if (/^\d{1,2}:\d{2}\s*(AM|PM)$/i.test(s)) return s.toUpperCase();
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var h = d.getHours(), m = d.getMinutes();
    var ap = h >= 12 ? 'PM' : 'AM';
    return (h % 12 || 12) + ':' + (m === 0 ? '00' : String(m).padStart(2,'0')) + ' ' + ap;
  }
  return s;
}

function schedIsBooked(name) {
  if (!name) return false;
  return /[a-zA-Z0-9]/.test(String(name).trim());
}

function schedToYMD(d) {
  return d.getFullYear() + '-' +
    String(d.getMonth()+1).padStart(2,'0') + '-' +
    String(d.getDate()).padStart(2,'0');
}

function schedThisMonday() {
  var t = new Date(); t.setHours(0,0,0,0);
  var day = t.getDay();
  t.setDate(t.getDate() + (day === 0 ? -6 : 1 - day));
  return t;
}

function schedFetch(callback) {
  fetch(SCHED_URL + '?action=schedule&_t=' + Date.now())
    .then(function(r){ return r.json(); })
    .then(function(json){
      if (!json || !json.rows) { callback(false); return; }
      var data = {};
      json.rows.forEach(function(row){
        var date   = (row.date || '').trim();
        var time   = schedNormalizeTime(row.time);
        var booked = schedIsBooked(row.customerName);
        if (!date || !time) return;
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
        if (!data[date]) data[date] = {};
        if (!data[date][time]) data[date][time] = { cap:0, booked:0, avail:0 };
        data[date][time].cap++;
        if (booked) data[date][time].booked++;
        data[date][time].avail = data[date][time].cap - data[date][time].booked;
      });
      schedData = data;
      callback(true);
    })
    .catch(function(){ callback(false); });
}

function schedShow() {
  document.getElementById('sched-loading').style.display = 'flex';
  document.getElementById('sched-picker').style.display  = 'none';
  document.getElementById('sched-error').style.display   = 'none';
  document.getElementById('sched-confirmed').style.display = 'none';
  schedWeekOff = 0;

  schedFetch(function(ok){
    document.getElementById('sched-loading').style.display = 'none';
    if (!ok) {
      document.getElementById('sched-error').style.display  = 'block';
      document.getElementById('sched-error').textContent    = '⚠ Could not load schedule.';
      return;
    }
    document.getElementById('sched-picker').style.display = 'block';
    schedRenderWeek();
  });
}

function schedRenderWeek() {
  var mon = schedThisMonday();
  mon.setDate(mon.getDate() + schedWeekOff * 7);
  var fri = new Date(mon); fri.setDate(mon.getDate() + 4);

  var MO = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  document.getElementById('sched-week-label').textContent =
    MO[mon.getMonth()] + ' ' + mon.getDate() + ' – ' +
    MO[fri.getMonth()] + ' ' + fri.getDate() + ', ' + fri.getFullYear();

  var DAYS = ['Mon','Tue','Wed','Thu','Fri'];
  var grid = document.getElementById('sched-day-grid');
  grid.innerHTML = '';
  var today = new Date(); today.setHours(0,0,0,0);

  for (var di = 0; di < 5; di++) {
    var day = new Date(mon); day.setDate(mon.getDate() + di);
    var key = schedToYMD(day);
    var isPast = day < today;
    var dd = schedData[key] || null;
    var totalAvail = dd ? SLOT_TIMES.reduce(function(s,t){ return s+(dd[t]?dd[t].avail:0); },0) : 0;

    var hdrCls = isPast || !dd ? '' : (totalAvail > 0 ? 'has-open' : 'all-full');
    var hdrCount = isPast ? 'Past' : (!dd ? 'No data' : (totalAvail > 0 ? totalAvail+' open' : 'Full'));

    var slotsHTML = SLOT_TIMES.map(function(t){
      var sd    = dd && dd[t];
      var avail = sd ? sd.avail : -1;
      var isChosen = selSlot && selSlot.date === key && selSlot.time === t;
      var cls, av;
      if (isPast)            { cls='past';   av='—'; }
      else if (!dd || !sd)   { cls='past';   av='—'; }
      else if (isChosen)     { cls='chosen'; av='✓'; }
      else if (avail <= 0)   { cls='full';   av='Full'; }
      else                   { cls='open';   av=avail+' left'; }
      var canClick = !isPast && sd && (avail > 0 || isChosen);
      var onclick  = canClick ? 'onclick="schedPickSlot(\''+key+'\',\''+t+'\')"' : '';
      return '<button class="sched-slot '+cls+'" '+onclick+'>'+
        '<span class="st">'+t.replace(':00','')+'</span>'+
        '<span class="sa">'+av+'</span>'+
        '</button>';
    }).join('');

    grid.innerHTML +=
      '<div class="sched-day">'+
        '<div class="sched-day-hdr '+hdrCls+'">'+
          '<span>'+DAYS[di]+' '+MO[day.getMonth()]+' '+day.getDate()+'</span>'+
          '<span class="sched-avail-count">'+hdrCount+'</span>'+
        '</div>'+
        '<div class="sched-slots">'+slotsHTML+'</div>'+
      '</div>';
  }
}

function schedShiftWeek(dir) {
  schedWeekOff += dir;
  if (schedWeekOff < 0) schedWeekOff = 0;
  schedRenderWeek();
}

function schedPickSlot(date, time) {
  selSlot = { date:date, time:time };
  document.getElementById('f-install-date').value = date;
  document.getElementById('f-install-time').value = time;
  calcPricing();

  var MO   = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var DAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var d    = new Date(date + 'T12:00:00');
  document.getElementById('sched-conf-date').textContent = DAYS[d.getDay()]+', '+MO[d.getMonth()]+' '+d.getDate()+', '+d.getFullYear();
  document.getElementById('sched-conf-time').textContent = '🕐 '+time;

  document.getElementById('sched-picker').style.display    = 'none';
  document.getElementById('sched-confirmed').style.display = 'flex';

  var mo = MO[d.getMonth()];
  document.getElementById('btn-mega').textContent = '⚡ Submit Mega — '+mo+' '+d.getDate()+' @ '+time;
  document.getElementById('btn-gig').textContent  = '🚀 Submit Gig — ' +mo+' '+d.getDate()+' @ '+time;
}

function schedClearSlot() {
  selSlot = null;
  document.getElementById('f-install-date').value = '';
  document.getElementById('f-install-time').value = '';
  document.getElementById('sched-confirmed').style.display = 'none';
  document.getElementById('sched-picker').style.display    = 'block';
  document.getElementById('proration-section').style.display = 'none';
  document.getElementById('btn-mega').textContent = '⚡ Submit — Mega Speed';
  document.getElementById('btn-gig').textContent  = '🚀 Submit — Gig Speed';
  schedRenderWeek();
}

function schedBookSlot(date, time, customerName, address) {
  fetch(SCHED_URL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ type:'booking', date:date, time:time, name:customerName, address:address })
  }).catch(function(){});

  if (schedData[date] && schedData[date][time]) {
    var s = schedData[date][time];
    if (s.avail > 0) { s.booked++; s.avail--; }
  }
}

var PKG = {
  mega: { base: 29.95, label: '$29.95' },
  gig:  { base: 39.95, label: '$39.95' }
};
var EERO = 5.00;
var PROC = 1.00;
var MODEM = 10.00;

function calcPricing() {
  if (!selPkg) return;
  var pkg = PKG[selPkg];
  document.getElementById('pr-internet').textContent = pkg.label;
  document.getElementById('pr-monthly').textContent  = '$' + (pkg.base + MODEM + EERO + PROC).toFixed(2);

  var dateEl = document.getElementById('f-install-date');
  var proSection = document.getElementById('proration-section');
  if (!dateEl.value) { proSection.style.display = 'none'; return; }

  var install = new Date(dateEl.value + 'T12:00:00');
  var nextFirst = new Date(install.getFullYear(), install.getMonth() + 1, 1);
  var diffDays = Math.round((nextFirst - install) / (1000 * 60 * 60 * 24));
  var daysInMonth = new Date(install.getFullYear(), install.getMonth() + 1, 0).getDate();
  var proratedInternet = (pkg.base / daysInMonth) * diffDays;
  var proratedEero     = (EERO / daysInMonth) * diffDays;
  var prorateToFirstBill = proratedInternet + proratedEero + PROC;

  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var day = install.getDate();

  document.getElementById('pr-prorate-label').textContent      = 'Internet (' + diffDays + ' days @ $' + (pkg.base / daysInMonth).toFixed(3) + '/day)';
  document.getElementById('pr-prorate-internet').textContent   = '$' + proratedInternet.toFixed(2);
  document.getElementById('pr-prorate-eero-label').textContent = 'eero 6+ (' + diffDays + ' days @ $' + (EERO / daysInMonth).toFixed(3) + '/day)';
  document.getElementById('pr-prorate-eero').textContent       = '$' + proratedEero.toFixed(2);
  document.getElementById('pr-prorate-total').textContent      = '$' + prorateToFirstBill.toFixed(2);
  var firstBillFeesOnly = MODEM + EERO + PROC;
  document.getElementById('pr-firstbill-total').textContent   = '$' + (firstBillFeesOnly + prorateToFirstBill).toFixed(2);
  document.getElementById('pr-firstbill-fees').textContent    = '$' + firstBillFeesOnly.toFixed(2);
  proSection.style.display = 'block';
}

function pickStatus(s) {
  selStatus = s;
  var map = {
    'Not Home':      { id:'sbt-nh',  cls:'act-nc' },
    'Brightspeed':   { id:'sbt-bs',  cls:'act-ni' },
    'In Contract':   { id:'sbt-ic',  cls:'act-vm' },
    'Not Interested':{ id:'sbt-ni',  cls:'act-ni' },
    'Go Back Later': { id:'sbt-gbl', cls:'act-cb' },
    'Vacant':        { id:'sbt-vac', cls:'act-nc' },
    'Business':      { id:'sbt-biz', cls:'act-vm' }
  };
  ['sbt-nh','sbt-bs','sbt-ic','sbt-ni','sbt-gbl','sbt-vac','sbt-biz'].forEach(function(sid) { document.getElementById(sid).className = 'stbtn'; });
  if (map[s]) document.getElementById(map[s].id).className = 'stbtn ' + map[s].cls;

  var needsNote = (s === 'Not Home' || s === 'Not Interested' || s === 'Go Back Later');
  var wrap = document.getElementById('ns-note-wrap');
  var note = document.getElementById('ns-note');
  if (wrap && note) {
    wrap.style.display = needsNote ? 'block' : 'none';
    if (!needsNote) note.value = '';
    if (needsNote) {
      note.placeholder =
        (s === 'Not Home') ? 'Example: will return after 5pm / left flyer' :
        (s === 'Go Back Later') ? 'Example: customer asked to come back Friday' :
        'Example: not interested — already has provider';
    }
  }
}

function fmtPhone(inp) {
  var v = inp.value.replace(/\D/g, '');
  if (v.length >= 10) v = '(' + v.slice(0,3) + ') ' + v.slice(3,6) + '-' + v.slice(6,10);
  inp.value = v;
}

function maybeWriteNewAddrToSheet(addr) {
  if (!addr._manuallyAdded) return;
  fetch(webhookURL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      type:      'add_address',
      address:   addr.address,
      city:      addr.city  || '',
      state:     addr.state || '',
      zip:       addr.zip   || '',
      lat:       addr.lat   != null ? addr.lat : '',
      lng:       addr.lng   != null ? addr.lng : '',
      pinDropped: addr._pinDropped ? true : false,
      addedBy:   repName,
      timestamp: new Date().toISOString()
    })
  }).catch(function(){});
}

// ──────────────────────────────────────────────────────────
//  SUBMIT
// ──────────────────────────────────────────────────────────
function getAddr() {
  for (var i = 0; i < addresses.length; i++) { if (addresses[i].id === activeId) return addresses[i]; }
  return null;
}

function submitSale(pkgLabel) {
  var addr = getAddr();
  if (!addr) { toast('No address selected', 't-err'); return; }

  var first   = document.getElementById('f-first').value.trim();
  var last    = document.getElementById('f-last').value.trim();
  var phone   = document.getElementById('f-phone').value.trim();
  var email   = document.getElementById('f-email').value.trim();
  var notes   = document.getElementById('f-notes').value.trim();
  var install = document.getElementById('f-install-date').value;

  if (!first || !last || !phone) {
    toast('⚠ Please fill in First Name, Last Name, and Phone', 't-err');
    return;
  }

  var pkg = PKG[selPkg];
  var monthlyTotal = (pkg.base + EERO + PROC).toFixed(2);
  var pricingSummary = pkgLabel + ' | Monthly: $' + monthlyTotal + ' | First Month: $16.00 (internet free)';
  if (install) {
    var installDate      = new Date(install + 'T12:00:00');
    var daysInMonth      = new Date(installDate.getFullYear(), installDate.getMonth() + 1, 0).getDate();
    var nextFirst        = new Date(installDate.getFullYear(), installDate.getMonth() + 1, 1);
    var diffDays         = Math.round((nextFirst - installDate) / (1000 * 60 * 60 * 24));
    var proratedInternet = (pkg.base / daysInMonth) * diffDays;
    var proratedEero     = (EERO / daysInMonth) * diffDays;
    var dueAtInstall     = (proratedInternet + proratedEero + PROC).toFixed(2);
    pricingSummary += ' | Estimated Proration: $' + dueAtInstall + ' (' + diffDays + ' day proration)';
  }

  var payload = {
    territory: (activeTerritory || ''),
    salesperson: repName,
    repPhone: repPhone,
    repEmail: repEmail,
    repWebsite: repWebsite,
    address: addr.address, city: addr.city||'', state: addr.state||'', zip: addr.zip||'',
    firstName: first, lastName: last, phone: phone, email: email,
    package: pricingSummary,
    installDate: selSlot ? selSlot.date : (install || ''),
    installTime: selSlot ? selSlot.time : '',
    notes: notes,
    status: 'Sale — ' + pkgLabel
  };

  sendData(payload);
  maybeWriteNewAddrToSheet(addr);

  if (selSlot) {
    var fullAddress = addr.address + (addr.city ? ', ' + addr.city : '') + (addr.state ? ', ' + addr.state : '');
    schedBookSlot(selSlot.date, selSlot.time, first + ' ' + last, fullAddress);
  }

  addr.status = (selPkg === 'mega') ? 'mega' : 'gig';
  addr.salesperson = repName;
  addr.sale   = { firstName: first, lastName: last, phone: phone, email: email, notes: notes };
  updateAddressStatus(addr, addr.status);
  addr.note = (notes || '').trim();
  if (addr.lat && addr.lng) placeMarker(addr);
  updateStats();
  sendHeartbeat();
  toast('✅ ' + pkgLabel + ' sold to ' + first + ' ' + last + '!', 't-ok');
  closeForm();
}

function submitStatus() {
  var addr = getAddr();
  if (!addr)      { toast('No address selected', 't-err'); return; }
  if (!selStatus) { toast('⚠ Pick a status first', 't-err'); return; }

  var nsWrap  = document.getElementById('ns-note-wrap');
  var nsNote  = document.getElementById('ns-note');
  var notes   = (nsWrap && nsWrap.style.display !== 'none' && nsNote)
    ? (nsNote.value || '').trim()
    : '';
  var payload = {
    salesperson: repName,
    address: addr.address, city: addr.city||'', state: addr.state||'', zip: addr.zip||'',
    firstName:'', lastName:'', phone:'', email:'',
    package:'', notes: notes,
    status: selStatus
  };

  sendData(payload);
  maybeWriteNewAddrToSheet(addr);

  var smap = { 'Not Home':'nothome','Brightspeed':'brightspeed','In Contract':'incontract','Not Interested':'notinterested','Go Back Later':'goback','Vacant':'vacant','Business':'business' };
  addr.status = smap[selStatus] || 'nocontact';
  addr.salesperson = repName;
  addr.note = notes || '';
  updateAddressStatus(addr, addr.status, notes);
  if (addr.lat && addr.lng) placeMarker(addr);
  updateStats();
  toast('📋 "' + selStatus + '" logged', 't-info');
  var nsNoteEl = document.getElementById('ns-note');
  if (nsNoteEl) nsNoteEl.value = '';
  closeForm();
}

function updateAddressStatus(addr, status, note) {
  // Always send — manual addresses have no sheetRow but the backend can still
  // log by address text, and we need the disposition to survive GPS refreshes.
  fetch(webhookURL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      type:            'status_update',
      territory:       (addr.territory || activeTerritory || ''),
      sheetRow:        addr.sheetRow || null,
      address:         addr.address  || '',
      city:            addr.city     || '',
      state:           addr.state    || '',
      zip:             addr.zip      || '',
      status:          status,
      salesperson:     repName,
      note:            (note || ''),
      dispositionNote: (note || '')
    })
  }).catch(function(){});
}

function sendData(payload) {
  if (!webhookURL) return;
  fetch(webhookURL, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type':'application/json' },
    body: JSON.stringify(payload)
  }).catch(function() {});
}

// ──────────────────────────────────────────────────────────
//  STATS
// ──────────────────────────────────────────────────────────
function updateStats() {
  document.getElementById('st-total').textContent = addresses.length;
  document.getElementById('st-sched').textContent = addresses.filter(function(a){ return a.status==='mega' || a.status==='gig'; }).length;
  document.getElementById('st-pend').textContent  = addresses.filter(function(a){
    var s = (a.status||'').toLowerCase();
    return !s || s === 'pending' || s === 'homes passed';
  }).length;
  // Show unique territory count in topbar for managers
  if (isManager && isManager()) {
    var territories = {};
    addresses.forEach(function(a){ if (a.territory) territories[a.territory] = true; });
    var tCount = Object.keys(territories).length;
    var stSched = document.getElementById('st-sched');
    if (stSched && tCount > 0) {
      stSched.parentElement.title = tCount + ' territories loaded';
    }
  }
}

// ──────────────────────────────────────────────────────────
//  MANAGER — Kasey Pelchy only
// ──────────────────────────────────────────────────────────
var MANAGER_NAMES  = ['kasey pelchy']; // ← add more names here, all lowercase
var heartbeatTimer = null;
var mgrAutoRefresh = null;

function isManager() {
  return MANAGER_NAMES.indexOf(repName.trim().toLowerCase()) >= 0;
}

function initManagerAccess() {
  if (isManager()) {
    document.getElementById('btn-manager').style.display = 'block';
  }
}

function sendHeartbeat(statusOverride) {
  var cleanName = (repName || '').trim();
  if (!webhookURL || isManager() || !cleanName || cleanName.toLowerCase() === 'rep') return;
  var status    = (statusOverride !== undefined) ? statusOverride : (repOnline ? 'online' : 'offline');
  var rn = cleanName.toLowerCase();
  var megaSales = addresses.filter(function(a){
    return a.status === 'mega' && ((a.salesperson || '').toLowerCase() === rn);
  }).length;
  var gigSales  = addresses.filter(function(a){
    return a.status === 'gig'  && ((a.salesperson || '').toLowerCase() === rn);
  }).length;
  var doorsWorked = addresses.filter(function(a){
    var st = String(a.status||'').toLowerCase();
    if (!st || st === 'pending') return false;
    return ((a.salesperson || '').toLowerCase() === rn);
  }).length;
  var firstSeen = '';
  try { firstSeen = localStorage.getItem('fieldos_session_start') || ''; } catch(e) {}
  fetch(webhookURL, {
    method: 'POST', mode: 'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      type:        'rep_heartbeat',
      salesperson: repName,
      status:      status,
      megaSales:   megaSales,
      gigSales:    gigSales,
      totalSales:   megaSales + gigSales,
      doorsWorked:  doorsWorked,
      firstSeen:    firstSeen,
      timestamp:    new Date().toISOString()
    })
  }).catch(function(){});
}

function startHeartbeat() {
  if (isManager()) return;
  sendHeartbeat();
  heartbeatTimer = setInterval(function() {
    if (repOnline) sendHeartbeat();
  }, 120000);
}

function stopHeartbeat() {
  clearInterval(heartbeatTimer);
  heartbeatTimer = null;
}

window.addEventListener('beforeunload', function() {
  if (!isManager()) sendHeartbeat('offline');
});

function openSignOutConfirm() {
  document.getElementById('signout-confirm').classList.add('open');
}
function closeSignOutConfirm() {
  document.getElementById('signout-confirm').classList.remove('open');
}
function confirmSignOut() {
  closeSignOutConfirm();
  sendHeartbeat('offline');
  stopHeartbeat();
  setTimeout(function() {
    repName   = 'Rep';
    try { localStorage.removeItem('fieldos_session_start'); } catch(e) {}
    repOnline = false;
    addresses = [];
    activeId  = null;
    selPkg    = null;
    selStatus = null;
    selSlot   = null;
    clearInterval(pollTimer);
    if (mapObj) { mapObj.remove(); mapObj = null; }

    document.getElementById('page-app').style.display   = 'none';
    document.getElementById('page-setup').style.display = 'flex';
    document.getElementById('rep-name').value = '';
    if (document.getElementById('rep-phone')) document.getElementById('rep-phone').value = '';
    if (document.getElementById('rep-email')) document.getElementById('rep-email').value = '';
    try { localStorage.removeItem('zito_rep_name'); localStorage.removeItem('zito_rep_phone'); localStorage.removeItem('zito_rep_email'); } catch(e) {}
    document.getElementById('launch-btn').disabled = true;
    document.getElementById('fetch-addr-status').textContent = '';
    document.getElementById('btn-manager').style.display = 'none';

    toast('👋 Signed out successfully', 't-info');
  }, 400);
}

function restoreRepProfile() {
  try {
    var n = localStorage.getItem('zito_rep_name')  || '';
    var p = localStorage.getItem('zito_rep_phone') || '';
    var e = localStorage.getItem('zito_rep_email') || '';
    if (n && document.getElementById('rep-name'))  document.getElementById('rep-name').value  = n;
    if (document.getElementById('rep-phone')) document.getElementById('rep-phone').value = p;
    if (document.getElementById('rep-email')) document.getElementById('rep-email').value = e;
  } catch(err) {}
}

window.addEventListener('load', function(){ try { restoreRepProfile(); } catch(e) {} });

function emailCustomerOffer(pkgKey) {
  var to = '';
  var custEmailEl = document.getElementById('f-email');
  if (custEmailEl) to = (custEmailEl.value || '').trim();
  if (!to) to = prompt('Customer email address to send the package info to:');
  if (!to) return;

  var rep = repName || 'Zito FieldOS';
  var rp  = repPhone || '';
  var re  = repEmail || '';
  var pkg = (pkgKey === 'gig') ? { name:'Gig Speed Fiber', speed:'1000/1000 Mbps', promo:'$49.95/mo', term:'2 years', reg:'$90.95/mo' }
                               : { name:'Mega Speed Fiber', speed:'400/400 Mbps',  promo:'$39.95/mo', term:'2 years', reg:'$87.39/mo' };

  var custFirst = '';
  var fn = document.getElementById('f-first');
  if (fn) custFirst = (fn.value || '').trim();

  var greet = custFirst ? ('Hi ' + custFirst + ',') : 'Hi there,';
  var subject = 'Zito Fiber Internet Package Details — ' + pkg.name;

  var bodyLines = [
    greet,
    '',
    'Here are the Zito Fiber details we discussed:',
    '',
    pkg.name,
    'Speed (Download/Upload): ' + pkg.speed,
    'Promo Price: ' + pkg.promo,
    'Promo Term: ' + pkg.term,
    'Regular Rate (after promo): ' + pkg.reg,
    '',
    'Whole‑Home Wi‑Fi (Required): eero 6+ mesh Wi‑Fi',
    '• $5/mo per eero device (coverage depends on home size)',
    '',
    'Ready to get started? Reply to this email and I can help schedule your install.',
    '',
    'Thanks,',
    rep + (rp ? (' | ' + rp) : ''),
    (re ? re : ''),
    repWebsite
  ];

  var mailto = 'mailto:' + encodeURIComponent(to)
    + '?subject=' + encodeURIComponent(subject)
    + '&body=' + encodeURIComponent(bodyLines.join('\n'));

  window.location.href = mailto;
}

function refreshMapMarkers() {
  if (!mapObj) return;

  // Clear existing markers (if any)
  if (window.mapMarkers) {
    Object.keys(mapMarkers).forEach(function(k){
      try { mapObj.removeLayer(mapMarkers[k]); } catch(e) {}
    });
  }
  window.mapMarkers = {};

  // Re-add markers for current addresses
  (addresses || []).forEach(function(a){
    if (a && a.lat != null && a.lng != null) {
      placeMarker(a);
    }
  });
}

function openManagerPanel() {
  document.getElementById('manager-modal').classList.add('open');
  refreshManagerPanel();
  mgrAutoRefresh = setInterval(refreshManagerPanel, 30000);
}
function closeManagerPanel() {
  document.getElementById('manager-modal').classList.remove('open');
  clearInterval(mgrAutoRefresh);
}
function refreshManagerPanel() {
  var btn = document.getElementById('mgr-refresh-btn');
  btn.classList.add('spinning');
  setTimeout(function(){ btn.classList.remove('spinning'); }, 500);

  fetch(webhookURL + '?action=repStatus&_t=' + Date.now())
    .then(function(r){ return r.json(); })
    .then(function(json){ renderRepList(json.reps || []); })
    .catch(function(){
      document.getElementById('mgr-rep-list').innerHTML =
        '<div class="mgr-empty"><div class="mgr-empty-icon">🔌</div>' +
        '<div class="mgr-empty-txt">Could not load rep data.<br>Make sure the Apps Script is deployed with the repStatus handler.</div></div>';
      updateMgrSummary(0, 0, 0);
      updateMgrPerformance({ doorsWorked:0,totalSales:0,megaSales:0,gigSales:0,onlineReps:0,activeHours:0 });
    });

  var now = new Date();
  document.getElementById('mgr-last-refresh').textContent =
    'Refreshed ' + now.toLocaleTimeString([], { hour:'2-digit', minute:'2-digit' });
}

function renderRepList(reps) {
  var onlineReps  = reps.filter(function(r){ return r.status === 'online'; });
  var offlineReps = reps.filter(function(r){ return r.status !== 'online'; });

  var megaTotal = reps.reduce(function(s,r){ return s + (Number(r.megaSales)||0); }, 0);
  var gigTotal  = reps.reduce(function(s,r){ return s + (Number(r.gigSales)||0); }, 0);
  var totalSales = reps.reduce(function(s,r){ return s + (Number(r.totalSales)||0); }, 0);

  var doorsWorked = reps.reduce(function(s,r){
    return s + (Number(r.doorsWorked)||0);
  }, 0);
  if (!doorsWorked && totalSales) doorsWorked = totalSales * 3;

  var nowMs = Date.now();
  var activeHours = onlineReps.reduce(function(s,r){
    var t0 = r.firstSeen ? new Date(r.firstSeen).getTime()
            : (r.lastSeen ? new Date(r.lastSeen).getTime() : nowMs);
    var hrs = Math.max((nowMs - t0) / 3600000, 0);
    return s + Math.max(hrs, 0.25);
  }, 0);

  updateMgrSummary(onlineReps.length, offlineReps.length, totalSales);
  updateMgrPerformance({
    doorsWorked: doorsWorked,
    totalSales: totalSales,
    megaSales: megaTotal,
    gigSales: gigTotal,
    onlineReps: onlineReps.length,
    activeHours: activeHours
  });

  if (reps.length === 0) {
    document.getElementById('mgr-rep-list').innerHTML =
      '<div class="mgr-empty"><div class="mgr-empty-icon">📡</div>' +
      '<div class="mgr-empty-txt">No reps have checked in yet.<br>Status updates appear here once reps log in.</div></div>';
    return;
  }

  var sorted = onlineReps.concat(offlineReps).sort(function(a,b){
    if (a.status==='online' && b.status!=='online') return -1;
    if (a.status!=='online' && b.status==='online') return  1;
    return (a.name||'').localeCompare(b.name||'');
  });

  document.getElementById('mgr-rep-list').innerHTML = sorted.map(function(rep) {
    var isOn    = rep.status === 'online';
    var parts   = (rep.name||'Rep').trim().split(/\s+/);
    var initials = parts.length >= 2
      ? (parts[0][0] + parts[parts.length-1][0]).toUpperCase()
      : (rep.name||'?').slice(0,2).toUpperCase();
    var lastSeen   = rep.lastSeen ? timeAgo(rep.lastSeen) : 'No activity';
    var mega       = Number(rep.megaSales)||0;
    var gig        = Number(rep.gigSales)||0;
    var total      = Number(rep.totalSales)||(mega+gig);
    var salesStr   = total + ' sale' + (total===1?'':'s');
    if (mega||gig) salesStr += ' (' + mega + ' Mega / ' + gig + ' Gig)';

    return '<div class="mgr-rep-card ' + (isOn?'rep-online':'rep-offline') + '">' +
      '<div class="mgr-rep-avatar">' + escHtml(initials) + '</div>' +
      '<div class="mgr-rep-info">' +
        '<div class="mgr-rep-name">' + escHtml(rep.name||'Unknown') + '</div>' +
        '<div class="mgr-rep-meta">Last seen: ' + lastSeen + '</div>' +
        (!isOn && rep.signOutTime ? '<div class="mgr-signout-time">Signed out ' + timeAgo(rep.signOutTime) + '</div>' : '') +
      '</div>' +
      '<div class="mgr-rep-right">' +
        '<div class="mgr-status-badge ' + (isOn?'online':'offline') + '">' +
          '<span class="mgr-status-dot"></span>' + (isOn?'ONLINE':'OFFLINE') +
        '</div>' +
        '<div class="mgr-rep-sales">' + escHtml(salesStr) + '</div>' +
      '</div>' +
    '</div>';
  }).join('');
}

function updateMgrSummary(online, offline, sales) {
  document.getElementById('mgr-count-online').textContent  = online;
  document.getElementById('mgr-count-offline').textContent = offline;
  document.getElementById('mgr-count-sales').textContent   = sales;
}

function updateMgrPerformance(metrics) {
  var doors   = Number(metrics.doorsWorked) || 0;
  var sales   = Number(metrics.totalSales)  || 0;
  var mega    = Number(metrics.megaSales)   || 0;
  var gig     = Number(metrics.gigSales)    || 0;
  var online  = Number(metrics.onlineReps)  || 0;
  var hours   = Number(metrics.activeHours) || 0;

  var closeRate = (doors > 0) ? (sales / doors) : 0;
  var pace = (hours > 0) ? (sales / hours) : 0;
  var denom = (mega + gig);
  var gigMix = (denom > 0) ? (gig / denom) : 0;
  var spr = (online > 0) ? (sales / online) : 0;

  function pct(x){ return Math.round(x * 100) + '%'; }
  function num1(x){ return (Math.round(x * 10) / 10).toFixed(1); }

  var elClose   = document.getElementById('mgr-m-close');
  var elPace    = document.getElementById('mgr-m-pace');
  var elGigMix  = document.getElementById('mgr-m-gigmix');
  var elSPR     = document.getElementById('mgr-m-spr');

  if (elClose)  elClose.textContent  = (doors > 0) ? pct(closeRate) : '—';
  if (elPace)   elPace.textContent   = (hours > 0) ? num1(pace) : '—';
  if (elGigMix) elGigMix.textContent = (denom > 0) ? pct(gigMix) : '—';
  if (elSPR)    elSPR.textContent    = (online > 0) ? num1(spr) : '—';

  var closeSub = document.getElementById('mgr-m-close-sub');
  var paceSub  = document.getElementById('mgr-m-pace-sub');
  var mixSub   = document.getElementById('mgr-m-gigmix-sub');
  var sprSub   = document.getElementById('mgr-m-spr-sub');

  if (closeSub) closeSub.textContent = (doors > 0) ? (sales + ' sales / ' + doors + ' worked') : 'No door activity reported';
  if (paceSub)  paceSub.textContent  = (hours > 0) ? ('Across ' + online + ' online rep' + (online===1?'':'s')) : '—';
  if (mixSub)   mixSub.textContent   = (denom > 0) ? (gig + ' Gig • ' + mega + ' Mega') : 'No sales reported';
  if (sprSub)   sprSub.textContent   = (online > 0) ? ('Online reps only') : '—';
}

function timeAgo(isoString) {
  var diff = Math.floor((Date.now() - new Date(isoString).getTime()) / 1000);
  if (diff < 60)   return diff + 's ago';
  if (diff < 3600) return Math.floor(diff/60) + 'm ago';
  return Math.floor(diff/3600) + 'h ago';
}
function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ──────────────────────────────────────────────────────────
//  BADGE & ONLINE/OFFLINE STATUS
// ──────────────────────────────────────────────────────────
var repOnline = false;

function initBadge() {
  repOnline = navigator.onLine;
  applyRepStatus();
  window.addEventListener('online',  function() { repOnline = true;  applyRepStatus(); sendHeartbeat('online');  });
  window.addEventListener('offline', function() { repOnline = false; applyRepStatus(); sendHeartbeat('offline'); });
  initManagerAccess();
  startHeartbeat();
}

function applyRepStatus() {
  var pill   = document.getElementById('tb-status-pill');
  var label  = document.getElementById('tb-status-label');
  var toggle = document.getElementById('badge-toggle');
  var btext  = document.getElementById('badge-status-text');

  if (repOnline) {
    pill.className   = 'is-online';
    label.textContent = 'ONLINE';
    toggle.className = 'badge-status-toggle online';
    btext.textContent = 'ONLINE';
  } else {
    pill.className   = 'is-offline';
    label.textContent = 'OFFLINE';
    toggle.className = 'badge-status-toggle offline';
    btext.textContent = 'OFFLINE';
  }
}

function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function popupHtmlForAddr(addr) {
  var status = (addr.status || 'pending').toString();
  var note   = (addr.note || '').toString().trim();

  return (
    '<div style="min-width:220px;">' +
      '<div style="font-weight:800;font-size:14px;margin-bottom:4px;">' + escapeHtml(addr.address || '') + '</div>' +
      '<div style="font-size:12px;opacity:.85;margin-bottom:8px;">' +
        escapeHtml([addr.city, addr.state, addr.zip].filter(Boolean).join(', ')) +
      '</div>' +
      '<div style="display:inline-block;padding:3px 8px;border-radius:999px;border:1px solid #ddd;font-size:12px;font-weight:700;margin-bottom:8px;">' +
        escapeHtml(status) +
      '</div>' +
      (note ? (
        '<div style="margin-top:8px;border-left:4px solid #46bba4;padding:6px 10px;background:rgba(70,187,164,0.08);border-radius:10px;">' +
          '<div style="font-size:11px;font-weight:800;letter-spacing:.02em;opacity:.85;margin-bottom:3px;">Disposition Note</div>' +
          '<div style="font-size:12px;">' + escapeHtml(note) + '</div>' +
        '</div>'
      ) : '') +
    '</div>'
  );
}

function toggleRepStatus() {
  repOnline = !repOnline;
  applyRepStatus();
  sendHeartbeat(repOnline ? 'online' : 'offline');
}

function openBadge() {
  var name = repName || 'Unknown Rep';
  document.getElementById('badge-rep-name').textContent = name;

  var p = repPhone || '';
  var e = repEmail || '';
  document.getElementById('badge-rep-phone').textContent = p ? p : '—';
  document.getElementById('badge-rep-email').textContent = e ? e : '—';
  document.getElementById('badge-rep-web').textContent   = repWebsite.replace(/^https?:\/\//,'');
  var phoneLink = document.getElementById('badge-phone-link');
  var emailLink = document.getElementById('badge-email-link');
  var webLink   = document.getElementById('badge-web-link');
  if (phoneLink) phoneLink.href = p ? ('tel:' + p.replace(/[^0-9+]/g,'')) : '#';
  if (emailLink) emailLink.href = e ? ('mailto:' + e) : '#';
  if (webLink)   webLink.href   = repWebsite;

  var parts    = name.trim().split(/\s+/);
  var initials = parts.length >= 2
    ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
    : name.slice(0, 2).toUpperCase();
  document.getElementById('badge-avatar').textContent = initials;

  var idNum = 0;
  for (var i = 0; i < name.length; i++) { idNum += name.charCodeAt(i); }
  document.getElementById('badge-id-num').textContent = 'REP-' + String(idNum).padStart(3, '0');

  var megaSales = addresses.filter(function(a){ return a.status === 'mega'; }).length;
  var gigSales  = addresses.filter(function(a){ return a.status === 'gig';  }).length;
  document.getElementById('badge-mega').textContent  = megaSales;
  document.getElementById('badge-gig').textContent   = gigSales;
  document.getElementById('badge-total').textContent = megaSales + gigSales;

  applyRepStatus();
  document.getElementById('badge-modal').classList.add('open');
}

function closeBadge() {
  document.getElementById('badge-modal').classList.remove('open');
}

var lastGPS = null;
var gpsWatchId = null;

function startGPSPing() {
  if (isManager()) return;       // managers see all territories — no GPS filtering
  if (!navigator.geolocation) return;

  // Watch position so it updates as they move
  gpsWatchId = navigator.geolocation.watchPosition(function(pos){
    lastGPS = {
      lat: pos.coords.latitude,
      lng: pos.coords.longitude,
      acc: pos.coords.accuracy || null,
      ts: Date.now()
    };

    // Send to server every update (or throttle if you want)
    pingNearbyAddresses();
  }, function(err){
    console.warn('Geolocation error:', err);
  }, {
    enableHighAccuracy: true,
    maximumAge: 10000,
    timeout: 10000
  });
}

function pingNearbyAddresses() {
  if (isManager()) return;       // managers don't filter by proximity
  if (!lastGPS) return;
  if (!repName) return; // your global repName after login/setup

  fetch(webhookURL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      type: 'rep_location',
      repName: repName,
      lat: lastGPS.lat,
      lng: lastGPS.lng,
      radiusMiles: 0.75, // adjust
      limit: 200
    })
  })
  .then(function(r){ return r.json(); })
  .then(function(json){
    if (!json || json.status !== 'ok') return;

    activeTerritory = (json.territory || '').trim();

    // ── Snapshot current rep-set dispositions before overwriting ──
    // This ensures that a sale/no-sale the rep already logged (on ANY address,
    // imported or manually added) is not wiped out by the next GPS refresh.
    var dispositionMap = {};
    addresses.forEach(function(a) {
      var s = (a.status || '').toLowerCase();
      // Only preserve statuses the rep actually set — not bare 'pending' from the sheet
      var REP_STATUSES = ['mega','gig','nothome','brightspeed','incontract','notinterested','goback','vacant','business'];
      if (REP_STATUSES.indexOf(s) >= 0) {
        var key = (a.address + '|' + (a.city || '')).toLowerCase().trim();
        dispositionMap[key] = { status: a.status, note: a.note || '', salesperson: a.salesperson || '', sale: a.sale || null };
      }
    });
    // Also hang on to manually added addresses so they survive the list rebuild
    var prevManual = addresses.filter(function(a) { return a._manuallyAdded; });

    // Build addresses list for UI
    addresses = (json.rows || []).map(function(row, i){
      var addr = {
        id: i,
        sheetRow: row.sheetRow,
        territory: row.territory,
        address: row.address,
        city: row.city,
        state: row.state,
        zip: row.zip,
        lat: row.lat != null ? parseFloat(row.lat) : null,
        lng: row.lng != null ? parseFloat(row.lng) : null,
        activeCount: (row.activeCount || '').toString().trim(),
        status: (row.status || 'pending').toLowerCase(),
        salesperson: (row.salesperson || '').trim(),
        note: (row.note || '').trim(),
        sale: null
      };
      // Restore any disposition the rep already logged for this address
      var key = (addr.address + '|' + (addr.city || '')).toLowerCase().trim();
      if (dispositionMap[key]) {
        addr.status     = dispositionMap[key].status;
        addr.note       = dispositionMap[key].note;
        addr.salesperson = dispositionMap[key].salesperson;
        addr.sale       = dispositionMap[key].sale;
      }
      return addr;
    });

    // Re-inject manually added addresses that didn't come back from the server
    var serverKeys = {};
    addresses.forEach(function(a) {
      serverKeys[(a.address + '|' + (a.city || '')).toLowerCase().trim()] = true;
    });
    prevManual.forEach(function(ma) {
      var key = (ma.address + '|' + (ma.city || '')).toLowerCase().trim();
      if (!serverKeys[key]) {
        var maxId = addresses.reduce(function(m, a) { return Math.max(m, a.id); }, -1);
        ma.id = maxId + 1;
        addresses.push(ma);
      }
    });

    updateStats();
    buildList();
    refreshMapMarkers();

    // Optional: keep map centered on rep
    if (mapObj) mapObj.setView([lastGPS.lat, lastGPS.lng], 17);
  })
  .catch(function(e){
    console.warn('rep_location failed', e);
  });
}

// ──────────────────────────────────────────────────────────
//  ADD ADDRESS MODAL
// ──────────────────────────────────────────────────────────
function openAddAddrModal() {
  ['new-addr-street','new-addr-city','new-addr-state','new-addr-zip'].forEach(function(id) {
    document.getElementById(id).value = '';
  });
  document.getElementById('btn-new-addr-submit').disabled = true;
  document.getElementById('add-addr-sending').style.display = 'none';
  document.getElementById('add-addr-modal').classList.add('open');
  setTimeout(function(){ document.getElementById('new-addr-street').focus(); }, 80);
}

function closeAddAddrModal() {
  document.getElementById('add-addr-modal').classList.remove('open');
}

function checkNewAddrReady() {
  var street = (document.getElementById('new-addr-street').value || '').trim();
  var city   = (document.getElementById('new-addr-city').value   || '').trim();
  document.getElementById('btn-new-addr-submit').disabled = !(street && city);
}

function submitNewAddress() {
  var street = (document.getElementById('new-addr-street').value || '').trim();
  var city   = (document.getElementById('new-addr-city').value   || '').trim();
  var state  = (document.getElementById('new-addr-state').value  || '').trim().toUpperCase();
  var zip    = (document.getElementById('new-addr-zip').value    || '').trim();

  if (!street || !city) {
    toast('⚠ Street address and city are required', 't-err');
    return;
  }

  var dup = addresses.find(function(a) {
    return a.address.toLowerCase() === street.toLowerCase() &&
           a.city.toLowerCase()    === city.toLowerCase();
  });
  if (dup) {
    toast('⚠ That address is already in the list', 't-err');
    return;
  }

  var newId = addresses.length > 0
    ? Math.max.apply(null, addresses.map(function(a){ return a.id; })) + 1
    : 0;

  var newAddr = {
    id:          newId,
    sheetRow:    null,
    address:     street,
    city:        city,
    state:       state,
    zip:         zip,
    lat:         null,
    lng:         null,
    activeCount: '',
    status:      'pending',
    salesperson: repName,
    sale:        null,
    _manuallyAdded: true
  };

  addresses.push(newAddr);
  updateStats();
  buildList();

  closeAddAddrModal();
  openForm(newId);

  var geocodeQuery = [street, city, state, zip].filter(Boolean).join(', ');
  var geocodeUrl = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=us&q=' + encodeURIComponent(geocodeQuery);
  fetch(geocodeUrl, { headers: { 'Accept': 'application/json', 'User-Agent': 'FieldSalesApp/1.0' } })
    .then(function(r){ return r.json(); })
    .then(function(data){
      if (data && data.length > 0) {
        newAddr.lat = parseFloat(data[0].lat);
        newAddr.lng = parseFloat(data[0].lon);
        if (mapObj) {
          placeMarker(newAddr);
          mapObj.panTo([newAddr.lat, newAddr.lng], { animate: true });
        }
      }
    })
    .catch(function(){});

  toast('📍 Address added — open the form to log a sale or no-sale', 't-info');
}

// ──────────────────────────────────────────────────────────
//  PIN DROP — tap the map to add a new address
// ──────────────────────────────────────────────────────────

function togglePinDropMode() {
  pinDropMode = !pinDropMode;
  var btn    = document.getElementById('btn-pin-drop');
  var banner = document.getElementById('pin-drop-banner');
  var mapEl  = document.getElementById('map');

  if (pinDropMode) {
    btn.classList.add('active');
    btn.textContent = '📍 Tap a Home…';
    banner.classList.add('show');
    mapEl.classList.add('pin-drop-mode');
    // Collapse sidebar on mobile so the full map is visible
    if (window.innerWidth <= 640 && sidebarOpen) toggleSidebar();
    toast('📍 Pin mode ON — tap any home on the map', 't-info');
  } else {
    cancelPinDropMode();
  }
}

function cancelPinDropMode() {
  pinDropMode = false;
  var btn    = document.getElementById('btn-pin-drop');
  var banner = document.getElementById('pin-drop-banner');
  var mapEl  = document.getElementById('map');
  if (btn)    { btn.classList.remove('active'); btn.textContent = '📍 Drop Pin'; }
  if (banner) { banner.classList.remove('show'); }
  if (mapEl)  { mapEl.classList.remove('pin-drop-mode'); }
  // Remove temp pin if still showing
  if (tempPinMarker && mapObj) { mapObj.removeLayer(tempPinMarker); tempPinMarker = null; }
}

function handleMapPinDrop(latlng) {
  // Immediately exit pin mode so accidental double-taps don't fire twice
  cancelPinDropMode();

  var lat = latlng.lat;
  var lng = latlng.lng;

  // Place a pulsing temp pin while we reverse-geocode
  var tempIcon = L.divIcon({
    className: '',
    html: '<div class="temp-pin-outer"><div class="temp-pin-inner"></div></div>',
    iconSize: [36, 36],
    iconAnchor: [18, 18]
  });
  tempPinMarker = L.marker([lat, lng], { icon: tempIcon }).addTo(mapObj);
  mapObj.panTo([lat, lng], { animate: true });
  toast('🔍 Looking up address…', 't-info');

  // Reverse geocode using Nominatim
  var url = 'https://nominatim.openstreetmap.org/reverse?format=json&lat=' +
            encodeURIComponent(lat) + '&lon=' + encodeURIComponent(lng) +
            '&zoom=18&addressdetails=1';

  fetch(url, { headers: { 'Accept': 'application/json', 'User-Agent': 'FieldSalesApp/1.0' } })
    .then(function(r) { return r.json(); })
    .then(function(data) {
      // Remove temp pin
      if (tempPinMarker && mapObj) { mapObj.removeLayer(tempPinMarker); tempPinMarker = null; }

      var a = data && data.address ? data.address : {};

      // Build street: house_number + road is the most reliable combo
      var street = ((a.house_number || '') + ' ' + (a.road || a.pedestrian || a.path || '')).trim();
      if (!street) {
        // Fall back to the display_name first segment, or coords
        street = data && data.display_name
          ? data.display_name.split(',')[0].trim()
          : ('Pin at ' + lat.toFixed(5) + ', ' + lng.toFixed(5));
      }

      var city  = a.city || a.town || a.village || a.hamlet || a.county || '';
      var state = a.state ? stateAbbr(a.state) : '';
      var zip   = a.postcode || '';

      addPinDropAddress(street, city, state, zip, lat, lng);
    })
    .catch(function() {
      if (tempPinMarker && mapObj) { mapObj.removeLayer(tempPinMarker); tempPinMarker = null; }
      // Still add with coords as fallback so the rep isn't left hanging
      var street = 'Pin at ' + lat.toFixed(5) + ', ' + lng.toFixed(5);
      addPinDropAddress(street, '', '', '', lat, lng);
      toast('⚠ Could not look up address — added as pin coordinates', 't-err');
    });
}

// Convert full US state name → 2-letter abbreviation
function stateAbbr(name) {
  var map = {
    'Alabama':'AL','Alaska':'AK','Arizona':'AZ','Arkansas':'AR','California':'CA',
    'Colorado':'CO','Connecticut':'CT','Delaware':'DE','Florida':'FL','Georgia':'GA',
    'Hawaii':'HI','Idaho':'ID','Illinois':'IL','Indiana':'IN','Iowa':'IA',
    'Kansas':'KS','Kentucky':'KY','Louisiana':'LA','Maine':'ME','Maryland':'MD',
    'Massachusetts':'MA','Michigan':'MI','Minnesota':'MN','Mississippi':'MS','Missouri':'MO',
    'Montana':'MT','Nebraska':'NE','Nevada':'NV','New Hampshire':'NH','New Jersey':'NJ',
    'New Mexico':'NM','New York':'NY','North Carolina':'NC','North Dakota':'ND','Ohio':'OH',
    'Oklahoma':'OK','Oregon':'OR','Pennsylvania':'PA','Rhode Island':'RI','South Carolina':'SC',
    'South Dakota':'SD','Tennessee':'TN','Texas':'TX','Utah':'UT','Vermont':'VT',
    'Virginia':'VA','Washington':'WA','West Virginia':'WV','Wisconsin':'WI','Wyoming':'WY'
  };
  return map[name] || name;
}

function addPinDropAddress(street, city, state, zip, lat, lng) {
  // Check for duplicate
  var dup = addresses.find(function(a) {
    return a.address.toLowerCase() === street.toLowerCase() &&
           (a.city || '').toLowerCase() === (city || '').toLowerCase();
  });
  if (dup) {
    toast('⚠ That address is already in the list', 't-err');
    openForm(dup.id);
    return;
  }

  var newId = addresses.length > 0
    ? Math.max.apply(null, addresses.map(function(a) { return a.id; })) + 1
    : 0;

  var newAddr = {
    id:             newId,
    sheetRow:       null,
    address:        street,
    city:           city,
    state:          state,
    zip:            zip,
    lat:            lat,
    lng:            lng,
    activeCount:    '',
    status:         'pending',
    salesperson:    repName,
    note:           '',
    sale:           null,
    _manuallyAdded: true,
    _pinDropped:    true
  };

  addresses.push(newAddr);
  updateStats();
  buildList();

  // Place the proper pending marker immediately (we already have coords)
  if (mapObj) placeMarker(newAddr);

  // Write to Google Sheet
  maybeWriteNewAddrToSheet(newAddr);

  // Open the sales form right away
  openForm(newId);

  toast('📍 ' + street + ' added!', 't-ok');
}

// ──────────────────────────────────────────────────────────
//  TOAST
// ──────────────────────────────────────────────────────────
function toast(msg, cls) {
  var el = document.getElementById('toast');
  el.textContent = msg;
  el.className   = cls + ' show';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function() { el.classList.remove('show'); }, 3200);
}

// Top Bar Drop Pin Hook
document.addEventListener('DOMContentLoaded', function() {
  var topDropBtn = document.getElementById('btn-drop-pin-top');
  if (topDropBtn && typeof enableDropMode === 'function') {
    topDropBtn.addEventListener('click', function() {
      enableDropMode();
    });
  }
});
