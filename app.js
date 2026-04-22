/**
 * ArtGrid — app.js
 * Milestone 1: Project scaffold with wired-up state and section visibility.
 * Milestones 2-5 will fill in the logic progressively.
 */

'use strict';

// ── State ─────────────────────────────────────────────────────────────────────
const state = {
  unit: 'cm',           // output unit — for all measurements & scale factor display
  canvasUnit: 'ft',     // input unit — how the artist knows their canvas size
  canvasWidth: null,    // in canvasUnit
  canvasHeight: null,
  canvasWidthCm: null,  // always stored internally in cm
  canvasHeightCm: null,
  image: null,          // HTMLImageElement
  imgNaturalW: 0,
  imgNaturalH: 0,
  fitMode: null,        // 'width' | 'height'
  scaleFactor: null,    // output units per image pixel
  imgOccupiedW: 0,      // output units
  imgOccupiedH: 0,      // output units
  imgOffsetX: 0,        // output units from canvas left
  imgOffsetY: 0,        // output units from canvas top
  marginLeft: 0,
  marginRight: 0,
  marginTop: 0,
  marginBottom: 0,
  placement: 'center',  // 'center' | 'top-left'
  gridMode: 'division',
  gridCols: 4,
  gridRows: 4,
  gridSpacingCol: null,
  gridSpacingRow: null,
  grayscale: false,
  pins: [],             // { id, label, imgX, imgY } — measurements computed on render
  pendingPoint: null,   // latest click measurement + guide metadata
  selectedPinId: null,
  uiZoom: 1,
  zoomControlsOpen: false,
  pointDetailsExpanded: false,
};

const UI_ZOOM_MIN = 1;
const UI_ZOOM_MAX = 4;
const UI_ZOOM_STEP = 0.25;

// ── DOM References ─────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const unitSelect        = $('unit-select');
const canvasUnitSelect  = $('canvas-unit-select');
const canvasWidthInput  = $('canvas-width');
const canvasHeightInput = $('canvas-height');
const placementSelect   = $('placement-select');
const imageUpload      = $('image-upload');
const fitInfoBox       = $('fit-info');
const sectionGrid      = $('section-grid');
const sectionWorkspace = $('section-workspace');
const sectionCanvasGrid= $('section-canvas-grid');
const gridModeRadios   = document.querySelectorAll('input[name="grid-mode"]');
const gridDivInputs    = $('grid-division-inputs');
const gridSpacInputs   = $('grid-spacing-inputs');
const gridColsInput    = $('grid-cols');
const gridRowsInput    = $('grid-rows');
const gridSpacColInput = $('grid-spacing-col');
const gridSpacRowInput = $('grid-spacing-row');
const btnDrawGrid      = $('btn-draw-grid');
const imageCanvas      = $('image-canvas');
const ctx              = imageCanvas.getContext('2d');
const lastPointInfo    = $('last-point-info');
const cellSizeInfo     = $('cell-size-info');
const setupSummary     = $('setup-summary');
const setupSummaryText = $('setup-summary-text');
const btnEditSetup     = $('btn-edit-setup');
const workspaceTopbar  = $('workspace-topbar');
const btnWorkspaceMenuToggle = $('btn-workspace-menu-toggle');
const imgPlacementRow  = $('img-placement-row');
const pinControls      = $('pin-controls');
const btnAddPin        = $('btn-add-pin');
const savedPointsMeta  = $('saved-points-meta');
const pinList          = $('pin-list');
const btnClearPins     = $('btn-clear-pins');
const canvasGridCanvas = $('canvas-grid-canvas');
const cgCtx            = canvasGridCanvas.getContext('2d');
const btnPrint         = $('btn-print');
const btnOpenSettings  = $('btn-open-settings');
const btnToggleGrayscale = $('btn-toggle-grayscale');
const btnToggleGridReference = $('btn-toggle-grid-reference');
const canvasWrapper    = $('canvas-wrapper');
const canvasZoomStage  = $('canvas-zoom-stage');
const imageStage       = $('image-stage');
const zoomControls     = $('zoom-controls');
const zoomLevel        = $('zoom-level');
const zoomRange        = $('zoom-range');
const btnZoomOut       = $('btn-zoom-out');
const btnZoomIn        = $('btn-zoom-in');
const btnZoomReset     = $('btn-zoom-reset');
const btnZoomToggle    = $('btn-zoom-toggle');
const currentPointSummary = $('current-point-summary');
const measurementPanel = $('measurement-panel');
const detailPanelResizer = $('detail-panel-resizer');

const TABLET_MIN_WIDTH = 601;
const TABLET_MAX_WIDTH = 900;
const TABLET_DETAIL_MIN_VH = 24;
const TABLET_DETAIL_MAX_VH = 58;

// ── Unit conversion helpers ────────────────────────────────────────
const TO_CM = { cm: 1, in: 2.54, ft: 30.48 };
const toCm   = (val, unit) => val * TO_CM[unit];
const fromCm = (val, unit) => val / TO_CM[unit];

// ── Unit Labels ──────────────────────────────────────────────
function updateUnitLabels() {
  // .unit-label = output unit (measurements)
  document.querySelectorAll('.unit-label').forEach(el => {
    el.textContent = state.unit;
  });
}

function updateCanvasUnitLabels() {
  // .canvas-unit-label = canvas input unit
  document.querySelectorAll('.canvas-unit-label').forEach(el => {
    el.textContent = state.canvasUnit;
  });
}

// ── Persistence ───────────────────────────────────────────────────────────────
function loadPreferences() {
  const savedUnit       = localStorage.getItem('artgrid_unit');
  const savedCanvasUnit = localStorage.getItem('artgrid_canvas_unit');
  const savedPlacement  = localStorage.getItem('artgrid_placement');
  const savedGrayscale  = localStorage.getItem('artgrid_grayscale');
  if (savedUnit) {
    state.unit = savedUnit;
    unitSelect.value = savedUnit;
  }
  if (savedCanvasUnit) {
    state.canvasUnit = savedCanvasUnit;
    canvasUnitSelect.value = savedCanvasUnit;
  }
  if (savedPlacement) {
    state.placement = savedPlacement;
    placementSelect.value = savedPlacement;
  }
  if (savedGrayscale) {
    state.grayscale = savedGrayscale === 'true';
  }
  updateUnitLabels();
  updateCanvasUnitLabels();
  syncGrayscaleButton();
}

function savePreferences() {
  localStorage.setItem('artgrid_unit',        state.unit);
  localStorage.setItem('artgrid_canvas_unit', state.canvasUnit);
  localStorage.setItem('artgrid_placement',   state.placement);
  localStorage.setItem('artgrid_grayscale',   state.grayscale);
}

function syncGrayscaleButton() {
  btnToggleGrayscale.textContent = state.grayscale ? 'Color View' : 'Grayscale';
  btnToggleGrayscale.setAttribute('aria-pressed', String(state.grayscale));
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function isCompactViewport() {
  return window.innerWidth <= 900;
}

function isTabletViewport() {
  return window.innerWidth >= TABLET_MIN_WIDTH && window.innerWidth <= TABLET_MAX_WIDTH;
}

function closeWorkspaceActionsMenu() {
  workspaceTopbar.classList.remove('actions-open');
  btnWorkspaceMenuToggle.setAttribute('aria-expanded', 'false');
}

function toggleWorkspaceActionsMenu() {
  const willOpen = !workspaceTopbar.classList.contains('actions-open');
  workspaceTopbar.classList.toggle('actions-open', willOpen);
  btnWorkspaceMenuToggle.setAttribute('aria-expanded', String(willOpen));
}

function clampVh(vh) {
  return clamp(vh, TABLET_DETAIL_MIN_VH, TABLET_DETAIL_MAX_VH);
}

function applyTabletDetailPanelHeight(vh) {
  if (!isTabletViewport() || !document.body.classList.contains('workspace-focus')) {
    measurementPanel.style.maxHeight = '';
    return;
  }
  measurementPanel.style.maxHeight = `${clampVh(vh)}vh`;
}

function setZoomControlsOpen(isOpen) {
  state.zoomControlsOpen = isOpen && !!state.image;
  zoomControls.classList.toggle('hidden', !state.zoomControlsOpen);
  btnZoomToggle.classList.toggle('hidden', !state.image);
  btnZoomToggle.setAttribute('aria-expanded', String(state.zoomControlsOpen));
}

function setPointDetailsExpanded(isExpanded) {
  const nextState = Boolean(isExpanded && state.pendingPoint);
  state.pointDetailsExpanded = nextState;
  lastPointInfo.classList.toggle('hidden', !nextState);
  currentPointSummary.setAttribute('aria-expanded', String(nextState));
}

function renderEmptyPointState(message = 'Select a point on the image') {
  currentPointSummary.classList.add('is-empty');
  currentPointSummary.disabled = true;
  currentPointSummary.textContent = message;
  lastPointInfo.classList.add('hidden');
  lastPointInfo.innerHTML = '';
  state.pointDetailsExpanded = false;
}

function renderPointSummary(point) {
  currentPointSummary.classList.remove('is-empty');
  currentPointSummary.disabled = false;
  currentPointSummary.innerHTML =
    `<span class="point-summary-main">` +
    `<span class="point-summary-cell">${point.cellName}</span>` +
    `<span>${point.xRefShort} / ${point.yRefShort}</span>` +
    `</span>` +
    `<span class="point-summary-metrics">` +
    `<span class="point-summary-value"><span class="legend-swatch swatch-blue"></span>${point.xDist} ${state.unit}</span>` +
    `<span class="point-summary-value"><span class="legend-swatch swatch-orange"></span>${point.yDist} ${state.unit}</span>` +
    `<span class="point-summary-action">${state.pointDetailsExpanded ? 'Hide details' : 'More details'}</span>` +
    `</span>`;
  currentPointSummary.setAttribute('aria-expanded', String(state.pointDetailsExpanded));
}

function updateZoomControlsUI() {
  zoomLevel.textContent = `Zoom ${Math.round(state.uiZoom * 100)}%`;
  zoomRange.value = String(Math.round(state.uiZoom * 100));
  btnZoomOut.disabled = state.uiZoom <= UI_ZOOM_MIN;
  btnZoomIn.disabled = state.uiZoom >= UI_ZOOM_MAX;
  btnZoomReset.disabled = state.uiZoom === 1;
}

function applyUiZoom(options = {}) {
  const { preserveViewportCenter = false } = options;
  if (!imageCanvas.width || !imageCanvas.height) {
    updateZoomControlsUI();
    return;
  }

  const prevZoom = imageCanvas._uiZoom || 1;
  const prevScaledW = imageCanvas.width * prevZoom;
  const prevScaledH = imageCanvas.height * prevZoom;
  const centerRatioX = prevScaledW > 0
    ? (canvasWrapper.scrollLeft + canvasWrapper.clientWidth / 2) / prevScaledW
    : 0.5;
  const centerRatioY = prevScaledH > 0
    ? (canvasWrapper.scrollTop + canvasWrapper.clientHeight / 2) / prevScaledH
    : 0.5;

  const scaledW = Math.max(1, Math.round(imageCanvas.width * state.uiZoom));
  const scaledH = Math.max(1, Math.round(imageCanvas.height * state.uiZoom));
  canvasZoomStage.style.width = `${scaledW}px`;
  canvasZoomStage.style.height = `${scaledH}px`;
  const zoomedIn = state.uiZoom > 1;
  canvasWrapper.style.overflowX = zoomedIn ? 'auto' : 'hidden';
  canvasWrapper.style.overflowY = zoomedIn ? 'auto' : 'hidden';
  imageCanvas._uiZoom = state.uiZoom;
  updateZoomControlsUI();

  if (preserveViewportCenter) {
    requestAnimationFrame(() => {
      const targetX = centerRatioX * scaledW - canvasWrapper.clientWidth / 2;
      const targetY = centerRatioY * scaledH - canvasWrapper.clientHeight / 2;
      canvasWrapper.scrollLeft = clamp(targetX, 0, Math.max(0, scaledW - canvasWrapper.clientWidth));
      canvasWrapper.scrollTop = clamp(targetY, 0, Math.max(0, scaledH - canvasWrapper.clientHeight));
    });
  } else if (!zoomedIn) {
    canvasWrapper.scrollLeft = 0;
    canvasWrapper.scrollTop = 0;
  }
}

function setUiZoom(nextZoom, options = {}) {
  const clamped = clamp(nextZoom, UI_ZOOM_MIN, UI_ZOOM_MAX);
  if (Math.abs(clamped - state.uiZoom) < 0.001) {
    updateZoomControlsUI();
    return;
  }
  state.uiZoom = clamped;
  applyUiZoom(options);
}

btnZoomOut.addEventListener('click', () => {
  setUiZoom(state.uiZoom - UI_ZOOM_STEP, { preserveViewportCenter: true });
});

btnZoomIn.addEventListener('click', () => {
  setUiZoom(state.uiZoom + UI_ZOOM_STEP, { preserveViewportCenter: true });
});

btnZoomReset.addEventListener('click', () => {
  setUiZoom(1, { preserveViewportCenter: true });
});

btnZoomToggle.addEventListener('click', e => {
  e.stopPropagation();
  setZoomControlsOpen(!state.zoomControlsOpen);
});

btnWorkspaceMenuToggle.addEventListener('click', e => {
  e.stopPropagation();
  toggleWorkspaceActionsMenu();
});

zoomControls.addEventListener('click', e => {
  e.stopPropagation();
});

zoomRange.addEventListener('input', () => {
  const nextZoom = parseInt(zoomRange.value, 10) / 100;
  setUiZoom(nextZoom, { preserveViewportCenter: true });
});

document.addEventListener('click', e => {
  if (!state.zoomControlsOpen) return;
  if (imageStage.contains(e.target)) return;
  setZoomControlsOpen(false);
});

document.addEventListener('click', e => {
  if (!workspaceTopbar.classList.contains('actions-open')) return;
  if (workspaceTopbar.contains(e.target)) return;
  closeWorkspaceActionsMenu();
});

currentPointSummary.addEventListener('click', () => {
  if (!state.pendingPoint) return;
  setPointDetailsExpanded(!state.pointDetailsExpanded);
  renderPointSummary(state.pendingPoint);
});

// ── Unit Change ────────────────────────────────────────────────────────────────
unitSelect.addEventListener('change', () => {
  state.unit = unitSelect.value;
  updateUnitLabels();
  savePreferences();
  // Recompute scale factor and refresh all measurements in new unit
  tryComputeFit();
  if (state.gridCols) {
    renderPinList();
    updateCellSizeInfo();
    renderCanvasGrid();
  }
});

canvasUnitSelect.addEventListener('change', () => {
  state.canvasUnit = canvasUnitSelect.value;
  updateCanvasUnitLabels();
  savePreferences();
  tryComputeFit();
  if (state.gridCols) {
    updateCellSizeInfo();
    renderPinList();
    renderCanvasGrid();
  }
});

placementSelect.addEventListener('change', () => {
  state.placement = placementSelect.value;
  savePreferences();
  tryComputeFit();
  if (state.gridCols) {
    updateCellSizeInfo();
    renderPinList();
    renderCanvasGrid();
  }
});

// ── Canvas + Image: trigger fit recalculation when all three are set ───────────
function tryComputeFit() {
  const w = parseFloat(canvasWidthInput.value);
  const h = parseFloat(canvasHeightInput.value);
  if (!state.image || !w || !h || w <= 0 || h <= 0) return;

  state.canvasWidth  = w;
  state.canvasHeight = h;

  // Convert canvas input unit → cm internally → output unit
  state.canvasWidthCm  = toCm(w, state.canvasUnit);
  state.canvasHeightCm = toCm(h, state.canvasUnit);
  const canvasWout = fromCm(state.canvasWidthCm,  state.unit);
  const canvasHout = fromCm(state.canvasHeightCm, state.unit);

  const canvasAspect = canvasWout / canvasHout;
  const imageAspect  = state.imgNaturalW / state.imgNaturalH;

  if (imageAspect > canvasAspect) {
    state.fitMode     = 'width';
    state.scaleFactor = canvasWout / state.imgNaturalW;
  } else {
    state.fitMode     = 'height';
    state.scaleFactor = canvasHout / state.imgNaturalH;
  }

  // Store occupied area and placement in output units
  state.imgOccupiedW = state.imgNaturalW * state.scaleFactor;
  state.imgOccupiedH = state.imgNaturalH * state.scaleFactor;

  if (state.placement === 'top-left') {
    state.imgOffsetX = 0;
    state.imgOffsetY = 0;
  } else {
    state.imgOffsetX = (canvasWout - state.imgOccupiedW) / 2;
    state.imgOffsetY = (canvasHout - state.imgOccupiedH) / 2;
  }

  state.marginLeft = Math.max(0, state.imgOffsetX);
  state.marginTop = Math.max(0, state.imgOffsetY);
  state.marginRight = Math.max(0, canvasWout - state.imgOffsetX - state.imgOccupiedW);
  state.marginBottom = Math.max(0, canvasHout - state.imgOffsetY - state.imgOccupiedH);

  fitInfoBox.innerHTML =
    `<strong>Image fit ready</strong>` +
    `<div class="stat-grid">` +
    `<span class="stat-pill"><span class="stat-label">Fit</span>${state.fitMode}</span>` +
    `<span class="stat-pill"><span class="stat-label">Area</span>${state.imgOccupiedW.toFixed(2)} × ${state.imgOccupiedH.toFixed(2)} ${state.unit}</span>` +
    `<span class="stat-pill"><span class="stat-label">Placement</span>${state.placement === 'center' ? 'Center' : 'Top-left'}</span>` +
    `<span class="stat-pill"><span class="stat-label">Margins</span>L ${state.marginLeft.toFixed(2)} · T ${state.marginTop.toFixed(2)} · R ${state.marginRight.toFixed(2)} · B ${state.marginBottom.toFixed(2)}</span>` +
    `</div>`;
  fitInfoBox.classList.remove('hidden');

  sectionGrid.classList.remove('hidden');
}

canvasWidthInput.addEventListener('input', tryComputeFit);
canvasHeightInput.addEventListener('input', tryComputeFit);

imageUpload.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const url = URL.createObjectURL(file);
  const img = new Image();
  img.onload = () => {
    state.image       = img;
    state.imgNaturalW = img.naturalWidth;
    state.imgNaturalH = img.naturalHeight;
    URL.revokeObjectURL(url);
    imgPlacementRow.classList.remove('hidden');
    tryComputeFit();
  };
  img.src = url;
});

// ── Grid Mode Toggle ───────────────────────────────────────────────────────────
gridModeRadios.forEach(radio => {
  radio.addEventListener('change', () => {
    state.gridMode = radio.value;
    if (state.gridMode === 'division') {
      gridDivInputs.classList.remove('hidden');
      gridSpacInputs.classList.add('hidden');
    } else {
      gridDivInputs.classList.add('hidden');
      gridSpacInputs.classList.remove('hidden');
    }
  });
});

// ── Draw Grid ──────────────────────────────────────────────────────────────────
btnDrawGrid.addEventListener('click', drawGrid);

function drawGrid() {
  if (!state.image || !state.scaleFactor) return;

  // Resolve grid cols/rows from inputs
  if (state.gridMode === 'division') {
    state.gridCols = Math.max(1, parseInt(gridColsInput.value) || 4);
    state.gridRows = Math.max(1, parseInt(gridRowsInput.value) || 4);
  } else {
    const sc = parseFloat(gridSpacColInput.value);
    const sr = parseFloat(gridSpacRowInput.value);
    if (!sc || !sr || sc <= 0 || sr <= 0) {
      alert('Please enter valid grid spacing values.');
      return;
    }
    // Convert spacing (canvas units) → pixel spacing on image
    state.gridSpacingCol = sc;
    state.gridSpacingRow = sr;
    const pxPerCol = sc / state.scaleFactor;
    const pxPerRow = sr / state.scaleFactor;
    state.gridCols = Math.floor(state.imgNaturalW / pxPerCol);
    state.gridRows = Math.floor(state.imgNaturalH / pxPerRow);
  }

  renderImageWithGrid();
  renderCanvasGrid();

  updateCellSizeInfo();

  sectionWorkspace.classList.remove('hidden');
  sectionCanvasGrid.classList.add('hidden');
  btnToggleGridReference.textContent = 'Show Grid Reference';
  document.body.classList.add('workspace-focus');
  setupSummary.classList.add('hidden');
  ['section-units', 'section-canvas', 'section-image', 'section-grid'].forEach(id => $(id).classList.add('hidden'));
  closeWorkspaceActionsMenu();
  applyTabletDetailPanelHeight(36);
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function updateCellSizeInfo() {
  if (!state.gridCols || !state.gridRows) return;
  const cellW = (state.imgOccupiedW / state.gridCols).toFixed(2);
  const cellH = (state.imgOccupiedH / state.gridRows).toFixed(2);

  cellSizeInfo.innerHTML =
    `<strong>${state.gridCols} × ${state.gridRows} grid</strong>` +
    `<div class="stat-grid">` +
    `<span class="stat-pill"><span class="stat-label">Cell</span>${cellW} × ${cellH} ${state.unit}</span>` +
    `<span class="stat-pill"><span class="stat-label">Start</span>${state.marginLeft.toFixed(2)} left · ${state.marginTop.toFixed(2)} top</span>` +
    `<span class="stat-pill"><span class="stat-label">Remaining</span>${state.marginRight.toFixed(2)} right · ${state.marginBottom.toFixed(2)} bottom</span>` +
    `</div>`;
  cellSizeInfo.classList.remove('hidden');
}

// ── Render Image + Canvas Preview (with dead space) ──────────────────────
function getMaxDisplayWidth() {
  if (document.body.classList.contains('workspace-focus')) {
    return Math.max(320, window.innerWidth - 64);
  }
  return Math.min(720, window.innerWidth - 48);
}

function getWrapperHeightLimit() {
  const wrapperStyle = window.getComputedStyle(canvasWrapper);
  const maxHeightPx = parseFloat(wrapperStyle.maxHeight);
  if (!Number.isFinite(maxHeightPx) || maxHeightPx <= 0) return null;

  const padTop = parseFloat(wrapperStyle.paddingTop) || 0;
  const padBottom = parseFloat(wrapperStyle.paddingBottom) || 0;
  const innerHeight = maxHeightPx - padTop - padBottom;
  return innerHeight > 0 ? innerHeight : null;
}

function renderImageWithGrid() {
  if (!state.canvasWidthCm || !state.image) return;

  const canvasWout = fromCm(state.canvasWidthCm, state.unit);
  const canvasHout = fromCm(state.canvasHeightCm, state.unit);

  const maxDispW   = getMaxDisplayWidth();
  const fallbackMaxDispH = Math.round(window.innerHeight * (isCompactViewport() ? 0.68 : 0.78));
  const wrapperHeightLimit = getWrapperHeightLimit();
  const maxDispH   = wrapperHeightLimit ? Math.min(fallbackMaxDispH, wrapperHeightLimit) : fallbackMaxDispH;
  const dispScale  = Math.min(maxDispW / canvasWout, maxDispH / canvasHout);

  const dispCanvasW = Math.round(canvasWout * dispScale);
  const dispCanvasH = Math.round(canvasHout * dispScale);
  const dispImgX    = Math.round(state.imgOffsetX * dispScale);
  const dispImgY    = Math.round(state.imgOffsetY * dispScale);
  const dispImgW    = Math.round(state.imgOccupiedW * dispScale);
  const dispImgH    = Math.round(state.imgOccupiedH * dispScale);
  const imgToDisp   = dispImgW / state.imgNaturalW;
  const imgToDispY  = dispImgH / state.imgNaturalH;

  imageCanvas.width  = dispCanvasW;
  imageCanvas.height = dispCanvasH;

  // Store for click/touch and pin mapping
  imageCanvas._dispImgX  = dispImgX;
  imageCanvas._dispImgY  = dispImgY;
  imageCanvas._dispImgW  = dispImgW;
  imageCanvas._dispImgH  = dispImgH;
  imageCanvas._imgToDispX = imgToDisp;
  imageCanvas._imgToDispY = imgToDispY;

  // Dead-space background (light grey hatched feel)
  ctx.fillStyle = '#d0d0d0';
  ctx.fillRect(0, 0, dispCanvasW, dispCanvasH);

  // Subtle crosshatch on dead space
  ctx.save();
  ctx.strokeStyle = 'rgba(0,0,0,0.07)';
  ctx.lineWidth = 1;
  for (let x = 0; x < dispCanvasW; x += 12) {
    ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, dispCanvasH); ctx.stroke();
  }
  for (let y = 0; y < dispCanvasH; y += 12) {
    ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(dispCanvasW, y); ctx.stroke();
  }
  ctx.restore();

  // Image drawn in its placement rectangle
  ctx.save();
  ctx.filter = state.grayscale ? 'grayscale(1)' : 'none';
  ctx.drawImage(state.image, dispImgX, dispImgY, dispImgW, dispImgH);
  ctx.restore();

  // Canvas outer border
  ctx.strokeStyle = '#555';
  ctx.lineWidth = 2;
  ctx.strokeRect(1, 1, dispCanvasW - 2, dispCanvasH - 2);

  // Image area dashed border
  ctx.save();
  ctx.strokeStyle = 'rgba(30, 100, 255, 0.55)';
  ctx.lineWidth = 1.5;
  ctx.setLineDash([5, 4]);
  ctx.strokeRect(dispImgX + 0.5, dispImgY + 0.5, dispImgW, dispImgH);
  ctx.restore();

  if (state.gridCols) {
    drawGridLinesInRect(ctx, dispImgX, dispImgY, dispImgW, dispImgH, state.gridCols, state.gridRows);
  }

  drawPendingPointGuides();

  redrawPins();

  btnZoomToggle.classList.remove('hidden');
  setZoomControlsOpen(state.zoomControlsOpen);
  applyUiZoom();
}

function drawPendingPointGuides() {
  if (!state.pendingPoint) return;

  const p = state.pendingPoint;
  const sx = imageCanvas._imgToDispX || 1;
  const sy = imageCanvas._imgToDispY || 1;
  const ox = imageCanvas._dispImgX || 0;
  const oy = imageCanvas._dispImgY || 0;

  const pointX = ox + p.imgX * sx;
  const pointY = oy + p.imgY * sy;

  const vBorderX = ox + p.vBorderPx * sx;
  const hBorderY = oy + p.hBorderPx * sy;

  ctx.save();
  ctx.setLineDash([6, 4]);
  ctx.lineWidth = 2;

  // Horizontal guide: nearest left/right border to clicked point
  ctx.strokeStyle = 'rgba(0, 145, 255, 0.95)';
  ctx.beginPath();
  ctx.moveTo(vBorderX, pointY);
  ctx.lineTo(pointX, pointY);
  ctx.stroke();

  // Vertical guide: nearest top/bottom border to clicked point
  ctx.strokeStyle = 'rgba(255, 150, 0, 0.95)';
  ctx.beginPath();
  ctx.moveTo(pointX, hBorderY);
  ctx.lineTo(pointX, pointY);
  ctx.stroke();

  ctx.setLineDash([]);

  // Anchor dots on referenced borders
  ctx.fillStyle = 'rgba(0, 145, 255, 0.95)';
  ctx.beginPath();
  ctx.arc(vBorderX, pointY, 3.5, 0, Math.PI * 2);
  ctx.fill();

  ctx.fillStyle = 'rgba(255, 150, 0, 0.95)';
  ctx.beginPath();
  ctx.arc(pointX, hBorderY, 3.5, 0, Math.PI * 2);
  ctx.fill();

  ctx.restore();
}

function drawGridLines(context, w, h, cols, rows) {
  drawGridLinesInRect(context, 0, 0, w, h, cols, rows);
}

function drawGridLinesInRect(context, x, y, w, h, cols, rows) {
  const colW = w / cols;
  const rowH = h / rows;

  context.save();
  context.strokeStyle = 'rgba(255, 50, 50, 0.75)';
  context.lineWidth   = 1;
  context.font        = `${Math.max(9, Math.min(12, colW * 0.18))}px sans-serif`;
  context.fillStyle   = 'rgba(255, 50, 50, 0.85)';

  // Vertical lines
  for (let c = 1; c < cols; c++) {
    context.beginPath();
    context.moveTo(x + c * colW, y);
    context.lineTo(x + c * colW, y + h);
    context.stroke();
  }
  // Horizontal lines
  for (let r = 1; r < rows; r++) {
    context.beginPath();
    context.moveTo(x, y + r * rowH);
    context.lineTo(x + w, y + r * rowH);
    context.stroke();
  }

  // Cell labels (A1, B2 …)
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const label = colLabel(c) + (r + 1);
      context.fillText(label, x + c * colW + 3, y + r * rowH + (context.font.match(/\d+/) ? parseInt(context.font) : 10) + 2);
    }
  }
  context.restore();
}

// A=0, B=1 … Z=25, AA=26 …
function colLabel(index) {
  let label = '';
  let n = index;
  do {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}

function buildPointMeasurement(imgX, imgY) {
  const colW = state.imgNaturalW / state.gridCols;
  const rowH = state.imgNaturalH / state.gridRows;
  const col = Math.min(Math.floor(imgX / colW), state.gridCols - 1);
  const row = Math.min(Math.floor(imgY / rowH), state.gridRows - 1);

  const cellName = colLabel(col) + (row + 1);
  const offsetLeftPx = imgX - col * colW;
  const offsetRightPx = colW - offsetLeftPx;
  const offsetTopPx = imgY - row * rowH;
  const offsetBottomPx = rowH - offsetTopPx;

  const useLeft = offsetLeftPx <= offsetRightPx;
  const useTop = offsetTopPx <= offsetBottomPx;

  const xDistPx = useLeft ? offsetLeftPx : offsetRightPx;
  const yDistPx = useTop ? offsetTopPx : offsetBottomPx;

  const rightNeighbor = col < state.gridCols - 1 ? `${colLabel(col + 1)}${row + 1}` : null;
  const bottomNeighbor = row < state.gridRows - 1 ? `${colLabel(col)}${row + 2}` : null;

  const xRefText = useLeft
    ? `left border of ${cellName}`
    : (rightNeighbor
      ? `right border of ${cellName} (left border of ${rightNeighbor})`
      : `right border of ${cellName}`);

  const yRefText = useTop
    ? `top border of ${cellName}`
    : (bottomNeighbor
      ? `bottom border of ${cellName} (top border of ${bottomNeighbor})`
      : `bottom border of ${cellName}`);

  const vBorderPx = col * colW + (useLeft ? 0 : colW);
  const hBorderPx = row * rowH + (useTop ? 0 : rowH);

  return {
    imgX,
    imgY,
    col,
    row,
    cellName,
    xDist: (xDistPx * state.scaleFactor).toFixed(2),
    yDist: (yDistPx * state.scaleFactor).toFixed(2),
    absX: (state.imgOffsetX + imgX * state.scaleFactor).toFixed(2),
    absY: (state.imgOffsetY + imgY * state.scaleFactor).toFixed(2),
    xRefText,
    yRefText,
    xRefShort: useLeft ? 'left' : 'right',
    yRefShort: useTop ? 'top' : 'bottom',
    vBorderPx,
    hBorderPx,
  };
}

function updateMeasurementInfo(point) {
  renderPointSummary(point);
  lastPointInfo.innerHTML =
    `<strong>Cell ${point.cellName}</strong><br>` +
    `<div class="measure-row"><span class="legend-swatch swatch-blue"></span><span class="measure-text">From <strong>${point.xRefText}</strong>: <strong>${point.xDist} ${state.unit}</strong></span></div>` +
    `<div class="measure-row"><span class="legend-swatch swatch-orange"></span><span class="measure-text">From <strong>${point.yRefText}</strong>: <strong>${point.yDist} ${state.unit}</strong></span></div>` +
    `<div class="stat-grid">` +
    `<span class="stat-pill"><span class="stat-label">Absolute</span>${point.absX} × ${point.absY} ${state.unit}</span>` +
    `<span class="stat-pill"><span class="stat-label">Saved</span>${state.pins.length}</span>` +
    `</div>`;
}

function showPendingPoint(point, options = {}) {
  const { selectedPinId = null } = options;
  state.pendingPoint = point;
  state.selectedPinId = selectedPinId;
  if (!selectedPinId) {
    state.pointDetailsExpanded = !isCompactViewport();
  }

  renderImageWithGrid();
  updateMeasurementInfo(point);
  setPointDetailsExpanded(state.pointDetailsExpanded);

  pinControls.classList.remove('hidden');
  renderPinList();
}

// ── Click / Touch to Measure ──────────────────────────────────────────────────
function measureAtViewportPoint(clientX, clientY) {
  if (!state.scaleFactor || !state.gridCols) return;

  const rect        = imageCanvas.getBoundingClientRect();
  const cssToCanvasX = imageCanvas.width / rect.width; // remain accurate under browser/UI zoom
  const cssToCanvasY = imageCanvas.height / rect.height;
  const canvasPxX    = (clientX - rect.left) * cssToCanvasX;
  const canvasPxY    = (clientY - rect.top) * cssToCanvasY;

  const dix = imageCanvas._dispImgX || 0;
  const diy = imageCanvas._dispImgY || 0;
  const diw = imageCanvas._dispImgW || imageCanvas.width;
  const dih = imageCanvas._dispImgH || imageCanvas.height;

  // Clip: ignore clicks on dead-space margins
  if (canvasPxX < dix || canvasPxX > dix + diw || canvasPxY < diy || canvasPxY > diy + dih) {
    lastPointInfo.innerHTML =
      `<span class="hint">Select a point inside the image area.</span>`;
    return;
  }

  const imgToDispX = imageCanvas._imgToDispX || 1;
  const imgToDispY = imageCanvas._imgToDispY || 1;
  const imgX = clamp((canvasPxX - dix) / imgToDispX, 0, state.imgNaturalW);
  const imgY = clamp((canvasPxY - diy) / imgToDispY, 0, state.imgNaturalH);

  setZoomControlsOpen(false);
  showPendingPoint(buildPointMeasurement(imgX, imgY));
}

imageCanvas.addEventListener('click', e => {
  measureAtViewportPoint(e.clientX, e.clientY);
});

imageCanvas.addEventListener('touchend', e => {
  e.preventDefault(); // prevent the ghost click that follows touch
  const touch = e.changedTouches[0];
  measureAtViewportPoint(touch.clientX, touch.clientY);
});

// ── Add Pin ───────────────────────────────────────────────────────────────────
btnAddPin.addEventListener('click', addPin);

function addPin() {
  if (!state.pendingPoint) return;
  const p     = state.pendingPoint;
  const suggestedLabel = `Point ${state.pins.length + 1}`;
  const enteredLabel = window.prompt('Point name (optional):', suggestedLabel);
  if (enteredLabel === null) return;
  const label = enteredLabel.trim() || suggestedLabel;
  // Store only raw pixel coords — measurements recomputed on render
  const pin   = { id: Date.now() + Math.floor(Math.random() * 1000), label, imgX: p.imgX, imgY: p.imgY };

  state.pins.push(pin);
  state.selectedPinId = pin.id;
  state.pendingPoint  = computePinMeasurements(pin);
  pinControls.classList.add('hidden');

  updateMeasurementInfo(state.pendingPoint);
  renderPinList();
  renderImageWithGrid(); // redraw to show pin marker
  btnClearPins.classList.remove('hidden');
}

// ── Render Pin Markers on Canvas ──────────────────────────────────────────────
function redrawPins() {
  const sx = imageCanvas._imgToDispX || 1;
  const sy = imageCanvas._imgToDispY || 1;
  const ox = imageCanvas._dispImgX   || 0;
  const oy = imageCanvas._dispImgY   || 0;
  state.pins.forEach(pin => {
    const dx = ox + pin.imgX * sx;
    const dy = oy + pin.imgY * sy;

    ctx.save();
    ctx.beginPath();
    ctx.arc(dx, dy, 5, 0, Math.PI * 2);
    ctx.fillStyle   = 'rgba(0,120,255,0.9)';
    ctx.fill();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth   = 1.5;
    ctx.stroke();
    ctx.font      = 'bold 11px sans-serif';
    ctx.fillStyle = '#003fa3';
    ctx.fillText(pin.label, dx + 7, dy - 4);
    ctx.restore();
  });
}

// ── Render Pin List ───────────────────────────────────────────────────────────
function computePinMeasurements(pin) {
  return buildPointMeasurement(pin.imgX, pin.imgY);
}

function renderPinList() {
  savedPointsMeta.textContent = `Saved points: ${state.pins.length}`;
  pinList.innerHTML = '';
  state.pins.forEach(pin => {
    const m  = computePinMeasurements(pin);
    const li = document.createElement('li');
    li.classList.toggle('active', pin.id === state.selectedPinId);
    li.tabIndex = 0;
    li.innerHTML =
      `<div class="pin-summary">` +
      `<span class="pin-name">${escapeHtml(pin.label)}</span>` +
      `<span class="pin-cell">${m.cellName}</span>` +
      `<span class="pin-metric"><span class="legend-swatch swatch-blue"></span>${m.xDist} ${state.unit}</span>` +
      `<span class="pin-metric"><span class="legend-swatch swatch-orange"></span>${m.yDist} ${state.unit}</span>` +
      `</div>` +
      `<div class="pin-meta hint">${m.xRefShort}/${m.yRefShort} · ${m.absX} × ${m.absY} ${state.unit}</div>` +
      `<button class="pin-remove" data-id="${pin.id}" title="Remove">✕</button>`;

    const selectPin = () => {
      showPendingPoint(m, { selectedPinId: pin.id });
    };

    li.addEventListener('click', selectPin);
    li.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        selectPin();
      }
    });

    pinList.appendChild(li);
  });

  pinList.querySelectorAll('.pin-remove').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const pinId = parseInt(btn.dataset.id, 10);
      state.pins = state.pins.filter(p => p.id !== parseInt(btn.dataset.id));
      if (state.selectedPinId === pinId) {
        state.selectedPinId = null;
        state.pendingPoint = null;
        pinControls.classList.add('hidden');
        renderEmptyPointState();
      }
      renderPinList();
      renderImageWithGrid();
      if (state.pins.length === 0) btnClearPins.classList.add('hidden');
    });
  });
}

// ── Clear All Pins ─────────────────────────────────────────────────────────────
btnClearPins.addEventListener('click', () => {
  state.pins = [];
  state.pendingPoint = null;
  state.selectedPinId = null;
  pinControls.classList.add('hidden');
  savedPointsMeta.textContent = 'Saved points: 0';
  renderEmptyPointState();
  renderPinList();
  renderImageWithGrid();
  btnClearPins.classList.add('hidden');
});
// ── Edit Setup ─────────────────────────────────────────────────────
function openSettingsView() {
  document.body.classList.remove('workspace-focus');
  closeWorkspaceActionsMenu();
  measurementPanel.style.maxHeight = '';
  sectionCanvasGrid.classList.add('hidden');
  btnToggleGridReference.textContent = 'Show Grid Reference';
  ['section-units', 'section-canvas', 'section-image', 'section-grid'].forEach(id => $(id).classList.remove('hidden'));
  if (state.image) imgPlacementRow.classList.remove('hidden'); // image already loaded
  setupSummary.classList.add('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

btnEditSetup.addEventListener('click', openSettingsView);
btnOpenSettings.addEventListener('click', openSettingsView);

btnToggleGrayscale.addEventListener('click', () => {
  state.grayscale = !state.grayscale;
  syncGrayscaleButton();
  savePreferences();
  if (state.image) {
    renderImageWithGrid();
  }
});

btnToggleGridReference.addEventListener('click', () => {
  const willShow = sectionCanvasGrid.classList.contains('hidden');
  sectionCanvasGrid.classList.toggle('hidden', !willShow);
  btnToggleGridReference.textContent = willShow ? 'Hide Grid Reference' : 'Show Grid Reference';

  if (willShow) {
    renderCanvasGrid();
    requestAnimationFrame(() => {
      sectionCanvasGrid.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }
});

detailPanelResizer.addEventListener('pointerdown', e => {
  if (!isTabletViewport() || !document.body.classList.contains('workspace-focus')) return;
  e.preventDefault();

  const startY = e.clientY;
  const currentHeightPx = measurementPanel.getBoundingClientRect().height;
  const startVh = (currentHeightPx / window.innerHeight) * 100;

  const onMove = moveEvent => {
    const deltaY = moveEvent.clientY - startY;
    const deltaVh = (deltaY / window.innerHeight) * 100;
    applyTabletDetailPanelHeight(startVh - deltaVh);
  };

  const onUp = () => {
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    window.removeEventListener('pointercancel', onUp);
  };

  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
  window.addEventListener('pointercancel', onUp);
});
// ── Render Canvas Grid (Companion View) ───────────────────────────────────────
const MAX_CANVAS_GRID_W = 700;
const MAX_CANVAS_GRID_H = 520;

function renderCanvasGrid() {
  const canvasW = fromCm(state.canvasWidthCm, state.unit);
  const canvasH = fromCm(state.canvasHeightCm, state.unit);
  const scale = Math.min(MAX_CANVAS_GRID_W / canvasW, MAX_CANVAS_GRID_H / canvasH);
  const dispW = Math.round(canvasW * scale);
  const dispH = Math.round(canvasH * scale);

  const imgX = Math.round(state.imgOffsetX * scale);
  const imgY = Math.round(state.imgOffsetY * scale);
  const imgW = Math.round(state.imgOccupiedW * scale);
  const imgH = Math.round(state.imgOccupiedH * scale);

  canvasGridCanvas.width  = dispW;
  canvasGridCanvas.height = dispH;

  // Background marks empty margins outside placed image area
  cgCtx.fillStyle = '#f2f2f2';
  cgCtx.fillRect(0, 0, dispW, dispH);

  cgCtx.fillStyle = '#ffffff';
  cgCtx.fillRect(imgX, imgY, imgW, imgH);

  // Outer border
  cgCtx.strokeStyle = '#333';
  cgCtx.lineWidth   = 2;
  cgCtx.strokeRect(1, 1, dispW - 2, dispH - 2);

  // Image area border
  cgCtx.strokeStyle = '#2f6fdd';
  cgCtx.lineWidth = 1.5;
  cgCtx.strokeRect(imgX, imgY, imgW, imgH);

  // Margin labels
  cgCtx.fillStyle = '#555';
  cgCtx.font = '12px sans-serif';
  cgCtx.fillText(`Margins (${state.unit})`, 8, 16);
  cgCtx.fillText(`L ${state.marginLeft.toFixed(2)}  R ${state.marginRight.toFixed(2)}  T ${state.marginTop.toFixed(2)}  B ${state.marginBottom.toFixed(2)}`, 8, 32);

  drawGridLinesInRect(cgCtx, imgX, imgY, imgW, imgH, state.gridCols, state.gridRows);
}

// ── Print ─────────────────────────────────────────────────────────────────────
btnPrint.addEventListener('click', () => window.print());

// ── Helpers ───────────────────────────────────────────────────────────────────
function escapeHtml(str) {
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Init ──────────────────────────────────────────────────────────────────────
loadPreferences();
renderEmptyPointState();

// Re-render on orientation change / resize (mobile)
let resizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    closeWorkspaceActionsMenu();
    if (state.pendingPoint) {
      state.pointDetailsExpanded = !isCompactViewport() && state.pointDetailsExpanded;
      renderPointSummary(state.pendingPoint);
      setPointDetailsExpanded(state.pointDetailsExpanded);
    }
    if (isTabletViewport() && document.body.classList.contains('workspace-focus')) {
      const currentMaxHeight = parseFloat(measurementPanel.style.maxHeight) || 36;
      applyTabletDetailPanelHeight(currentMaxHeight);
    } else {
      measurementPanel.style.maxHeight = '';
    }
    if (state.image && state.gridCols) {
      tryComputeFit();        // recalculate scale factor for new output unit layout
      renderImageWithGrid();
      updateCellSizeInfo();
      renderCanvasGrid();
    }
  }, 200);
});

updateZoomControlsUI();
