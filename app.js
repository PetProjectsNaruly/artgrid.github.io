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
};

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
const imgPlacementRow  = $('img-placement-row');
const pinControls      = $('pin-controls');
const pinLabelInput    = $('pin-label-input');
const btnAddPin        = $('btn-add-pin');
const pinList          = $('pin-list');
const btnClearPins     = $('btn-clear-pins');
const canvasGridCanvas = $('canvas-grid-canvas');
const cgCtx            = canvasGridCanvas.getContext('2d');
const btnPrint         = $('btn-print');
const btnOpenSettings  = $('btn-open-settings');
const btnToggleGrayscale = $('btn-toggle-grayscale');
const btnToggleGridReference = $('btn-toggle-grid-reference');

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

function renderImageWithGrid() {
  if (!state.canvasWidthCm || !state.image) return;

  const canvasWout = fromCm(state.canvasWidthCm, state.unit);
  const canvasHout = fromCm(state.canvasHeightCm, state.unit);

  const maxDispW   = getMaxDisplayWidth();
  const maxDispH   = Math.round(window.innerHeight * 0.78);
  const dispScale  = Math.min(maxDispW / canvasWout, maxDispH / canvasHout);

  const dispCanvasW = Math.round(canvasWout * dispScale);
  const dispCanvasH = Math.round(canvasHout * dispScale);
  const dispImgX    = Math.round(state.imgOffsetX * dispScale);
  const dispImgY    = Math.round(state.imgOffsetY * dispScale);
  const dispImgW    = Math.round(state.imgOccupiedW * dispScale);
  const dispImgH    = Math.round(state.imgOccupiedH * dispScale);
  const imgToDisp   = dispImgW / state.imgNaturalW;

  imageCanvas.width  = dispCanvasW;
  imageCanvas.height = dispCanvasH;

  // Store for click/touch and pin mapping
  imageCanvas._dispImgX  = dispImgX;
  imageCanvas._dispImgY  = dispImgY;
  imageCanvas._dispImgW  = dispImgW;
  imageCanvas._dispImgH  = dispImgH;
  imageCanvas._imgToDisp = imgToDisp;

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
}

function drawPendingPointGuides() {
  if (!state.pendingPoint) return;

  const p = state.pendingPoint;
  const s = imageCanvas._imgToDisp || 1;
  const ox = imageCanvas._dispImgX || 0;
  const oy = imageCanvas._dispImgY || 0;

  const pointX = ox + p.imgX * s;
  const pointY = oy + p.imgY * s;

  const vBorderX = ox + p.vBorderPx * s;
  const hBorderY = oy + p.hBorderPx * s;

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
    vBorderPx,
    hBorderPx,
  };
}

function updateMeasurementInfo(point) {
  lastPointInfo.innerHTML =
    `<strong>Cell ${point.cellName}</strong><br>` +
    `<div class="measure-row"><span class="legend-swatch swatch-blue"></span><span class="measure-text">From <strong>${point.xRefText}</strong>: <strong>${point.xDist} ${state.unit}</strong></span></div>` +
    `<div class="measure-row"><span class="legend-swatch swatch-orange"></span><span class="measure-text">From <strong>${point.yRefText}</strong>: <strong>${point.yDist} ${state.unit}</strong></span></div>` +
    `<div class="stat-grid"><span class="stat-pill"><span class="stat-label">Absolute</span>${point.absX} × ${point.absY} ${state.unit}</span></div>`;
}

function showPendingPoint(point, options = {}) {
  const { selectedPinId = null } = options;
  state.pendingPoint = point;
  state.selectedPinId = selectedPinId;

  renderImageWithGrid();
  updateMeasurementInfo(point);

  pinControls.classList.remove('hidden');
  renderPinList();
}

// ── Click / Touch to Measure ──────────────────────────────────────────────────
function measureAtViewportPoint(clientX, clientY) {
  if (!state.scaleFactor || !state.gridCols) return;

  const rect        = imageCanvas.getBoundingClientRect();
  const cssToCanvas = imageCanvas.width / rect.width; // handle any CSS scaling
  const canvasPxX   = (clientX - rect.left) * cssToCanvas;
  const canvasPxY   = (clientY - rect.top)  * cssToCanvas;

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

  const imgToDisp = imageCanvas._imgToDisp || 1;
  const imgX = (canvasPxX - dix) / imgToDisp;
  const imgY = (canvasPxY - diy) / imgToDisp;

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
pinLabelInput.addEventListener('keydown', e => { if (e.key === 'Enter') addPin(); });

function addPin() {
  if (!state.pendingPoint) return;
  const p     = state.pendingPoint;
  const label = pinLabelInput.value.trim() || `Point ${state.pins.length + 1}`;
  // Store only raw pixel coords — measurements recomputed on render
  const pin   = { id: Date.now(), label, imgX: p.imgX, imgY: p.imgY };

  state.pins.push(pin);
  pinLabelInput.value = '';
  state.pendingPoint  = null;
  state.selectedPinId = null;
  pinControls.classList.add('hidden');

  renderPinList();
  renderImageWithGrid(); // redraw to show pin marker
  btnClearPins.classList.remove('hidden');
}

// ── Render Pin Markers on Canvas ──────────────────────────────────────────────
function redrawPins() {
  const s  = imageCanvas._imgToDisp  || 1;
  const ox = imageCanvas._dispImgX   || 0;
  const oy = imageCanvas._dispImgY   || 0;
  state.pins.forEach(pin => {
    const dx = ox + pin.imgX * s;
    const dy = oy + pin.imgY * s;

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
  pinList.innerHTML = '';
  state.pins.forEach(pin => {
    const m  = computePinMeasurements(pin);
    const li = document.createElement('li');
    li.classList.toggle('active', pin.id === state.selectedPinId);
    li.tabIndex = 0;
    li.innerHTML =
      `<span class="pin-name">${escapeHtml(pin.label)}</span>` +
      `Cell <strong>${m.cellName}</strong><br>` +
      `From ${m.xRefText}: <strong>${m.xDist} ${state.unit}</strong><br>` +
      `From ${m.yRefText}: <strong>${m.yDist} ${state.unit}</strong><br>` +
      `<span class="hint">${m.absX} × ${m.absY} ${state.unit}</span>` +
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
        lastPointInfo.innerHTML = 'Select a point on the image.';
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
  lastPointInfo.innerHTML = 'Select a point on the image.';
  renderPinList();
  renderImageWithGrid();
  btnClearPins.classList.add('hidden');
});
// ── Edit Setup ─────────────────────────────────────────────────────
function openSettingsView() {
  document.body.classList.remove('workspace-focus');
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

// Re-render on orientation change / resize (mobile)
let resizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    if (state.image && state.gridCols) {
      tryComputeFit();        // recalculate scale factor for new output unit layout
      renderImageWithGrid();
      updateCellSizeInfo();
      renderCanvasGrid();
    }
  }, 200);
});
