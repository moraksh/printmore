/**
 * designer.js — Designer class for PrintMore
 */

'use strict';

const MM_TO_PX = 3.7795;
const GRID_SIZE_MM = 5; // grid every 5mm
const MIN_ELEMENT_SIZE = 5; // mm

class Designer {
  constructor(layoutId) {
    this.layoutId = layoutId;
      this.layout = null;
      this.elements = [];
      this.selectedId = null;
      this.selectedIds = [];
      this.activeTool = 'select';
    this.zoom = 1;
    this.showGrid = true;

    // Drag/resize state
    this.dragState = null;
      this.resizeState = null;
      this.drawState = null;    // for drawing new elements by drag
      this.marqueeState = null; // for left-drag multi-select
      this.suppressNextContextMenu = false;

    // Field drag from panel
    this.fieldDragName = null;

    // Column / row resize state
    this.colResizeState = null;
    this.rowResizeState = null;

    // Selected table cell
    this.selectedCell = null;

    // Default style for new elements (persists across element creation)
    this.defaultStyle = {
      fontSize: 12,
      fontFamily: 'Arial',
      fontWeight: 'normal',
      fontStyle: 'normal',
      textDecoration: 'none',
      color: '#000000',
      textAlign: 'left',
    };

    // History
    this.history = [];
    this.historyIndex = -1;

    // DOM refs
    this.pageCanvas = null;
    this.gridCanvas = null;
    this.canvasOuter = null;
    this.scrollArea = null;

    this._boundMouseMove = this._onMouseMove.bind(this);
    this._boundMouseUp = this._onMouseUp.bind(this);
    this._boundKeyDown = this._onKeyDown.bind(this);
    this._boundRulerScroll = null;
    this._debounceSaveTimer = null;
  }

  // ===== Init =====
  init() {
    this.layout = window.getLayoutById(this.layoutId);
    if (!this.layout) return;
    this.layout.texts = this.layout.texts || [];
    this.layout.defaultStyle = this.layout.defaultStyle || {};
    this.defaultStyle = { ...this.defaultStyle, ...this.layout.defaultStyle };

    this.elements = JSON.parse(JSON.stringify(this.layout.elements || []));

    this.pageCanvas = document.getElementById('page-canvas');
    this.gridCanvas = document.getElementById('grid-canvas');
    this.canvasOuter = document.getElementById('canvas-outer');
    this.scrollArea = document.getElementById('canvas-scroll-area');
    if (!this._boundRulerScroll) {
      this._boundRulerScroll = this._syncRulersWithScroll.bind(this);
      this.scrollArea.addEventListener('scroll', this._boundRulerScroll);
    }

    this.zoom = 1;
    const zoomLabelEl = document.getElementById('zoom-label');
    if (zoomLabelEl) zoomLabelEl.textContent = '100%';

    this.initCanvas();
    this.renderElements();
    this.renderFieldsList();
    this.initToolbarEvents();
    this.initCanvasEvents();
    this._showNoSelectionPanel(); // populate inline page settings on load
    this.saveToHistory();
    this.updateUndoRedo();

    document.addEventListener('mousemove', this._boundMouseMove);
    document.addEventListener('mouseup', this._boundMouseUp);
    document.addEventListener('keydown', this._boundKeyDown);
  }

  destroy() {
    document.removeEventListener('mousemove', this._boundMouseMove);
    document.removeEventListener('mouseup', this._boundMouseUp);
    document.removeEventListener('keydown', this._boundKeyDown);
    if (this.scrollArea && this._boundRulerScroll) {
      this.scrollArea.removeEventListener('scroll', this._boundRulerScroll);
      this._boundRulerScroll = null;
    }
    // Clean up context menu
    this._hideContextMenu();
  }

  // ===== Canvas Setup =====
  initCanvas() {
    const { pageWidthPx, pageHeightPx } = this._getPageDimensions();
    this.pageCanvas.style.width = pageWidthPx + 'px';
    this.pageCanvas.style.height = pageHeightPx + 'px';
    this.canvasOuter.style.minWidth = (pageWidthPx + 80) + 'px';
    this.canvasOuter.style.minHeight = (pageHeightPx + 80) + 'px';

    // Grid canvas
    this.gridCanvas.width = pageWidthPx;
    this.gridCanvas.height = pageHeightPx;
    this.gridCanvas.style.width = pageWidthPx + 'px';
    this.gridCanvas.style.height = pageHeightPx + 'px';

    this._drawGrid();
    this._drawMarginGuide();
    this._drawRulers();
    this._syncRulersWithScroll();
    this._drawZoneOverlays();
  }

  _drawZoneOverlays() {
    this.pageCanvas.querySelectorAll('.zone-overlay').forEach(el => el.remove());
    const p = this.layout.page;
    const { pageHeightPx } = this._getPageDimensions();

    if (p.headerEnabled && p.headerHeight > 0) {
      const hPx = p.headerHeight * MM_TO_PX * this.zoom;
      const overlay = document.createElement('div');
      overlay.className = 'zone-overlay zone-overlay-header';
      overlay.style.height = hPx + 'px';
      const lbl = document.createElement('span');
      lbl.className = 'zone-label zone-label-header';
      lbl.textContent = 'HEADER' + (p.headerPages === 'first' ? ' (1st page)' : '');
      overlay.appendChild(lbl);
      this.pageCanvas.appendChild(overlay);
    }

    if (p.footerEnabled && p.footerHeight > 0) {
      const fPx = p.footerHeight * MM_TO_PX * this.zoom;
      const overlay = document.createElement('div');
      overlay.className = 'zone-overlay zone-overlay-footer';
      overlay.style.cssText += `height:${fPx}px;`;
      const lbl = document.createElement('span');
      lbl.className = 'zone-label zone-label-footer';
      lbl.textContent = 'FOOTER' + (p.footerPages === 'last' ? ' (last page)' : '');
      overlay.appendChild(lbl);
      this.pageCanvas.appendChild(overlay);
    }
  }

  _getPageDimensions() {
    const page = this.layout.page;
    const sizes = window.PAGE_SIZES;
    let width;
    let height;
    if (page.size === 'custom') {
      width = Math.max(20, parseFloat(page.customWidthMm ?? page.customWidth) || 210);
      height = Math.max(20, parseFloat(page.customHeightMm ?? page.customHeight) || 297);
    } else {
      ({ width, height } = sizes[page.size] || sizes['A4']);
    }
    if (page.orientation === 'landscape') { [width, height] = [height, width]; }
    return {
      pageWidthMm: width,
      pageHeightMm: height,
      pageWidthPx: width * MM_TO_PX * this.zoom,
      pageHeightPx: height * MM_TO_PX * this.zoom,
    };
  }

  _drawGrid() {
    const ctx = this.gridCanvas.getContext('2d');
    const { pageWidthPx, pageHeightPx } = this._getPageDimensions();
    ctx.clearRect(0, 0, pageWidthPx, pageHeightPx);
    if (!this.showGrid) return;

    const stepPx = GRID_SIZE_MM * MM_TO_PX * this.zoom;
    ctx.strokeStyle = 'rgba(180,180,220,0.2)';
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    for (let x = 0; x <= pageWidthPx; x += stepPx) {
      ctx.moveTo(x, 0);
      ctx.lineTo(x, pageHeightPx);
    }
    for (let y = 0; y <= pageHeightPx; y += stepPx) {
      ctx.moveTo(0, y);
      ctx.lineTo(pageWidthPx, y);
    }
    ctx.stroke();
  }

  _drawMarginGuide() {
    // Remove old guides
    this.pageCanvas.querySelectorAll('.margin-guide').forEach(el => el.remove());
    const p = this.layout.page;
    const { pageWidthPx, pageHeightPx } = this._getPageDimensions();
    const top = p.marginTop * MM_TO_PX * this.zoom;
    const left = p.marginLeft * MM_TO_PX * this.zoom;
    const right = p.marginRight * MM_TO_PX * this.zoom;
    const bottom = p.marginBottom * MM_TO_PX * this.zoom;
    const guide = document.createElement('div');
    guide.className = 'margin-guide';
    guide.style.top = top + 'px';
    guide.style.left = left + 'px';
    guide.style.width = (pageWidthPx - left - right) + 'px';
    guide.style.height = (pageHeightPx - top - bottom) + 'px';
    this.pageCanvas.appendChild(guide);
  }

  _drawRulers() {
    const rulerH = document.getElementById('ruler-h');
    const rulerV = document.getElementById('ruler-v');
    const { pageWidthMm, pageHeightMm, pageWidthPx, pageHeightPx } = this._getPageDimensions();
    rulerH.innerHTML = '';
    rulerV.innerHTML = '';

    // Horizontal
    const canvasH = document.createElement('canvas');
    canvasH.width = pageWidthPx + 80;
    canvasH.height = 20;
    canvasH.style.position = 'absolute';
    canvasH.style.left = '0';
    canvasH.style.top = '0';
    const ctxH = canvasH.getContext('2d');
    ctxH.fillStyle = '#606080';
    ctxH.font = '9px Arial';
    ctxH.textBaseline = 'bottom';
    const hStep = MM_TO_PX * this.zoom;
    for (let mm = 0; mm <= pageWidthMm; mm += 10) {
      const x = 40 + mm * hStep;
      ctxH.beginPath();
      ctxH.strokeStyle = '#606080';
      ctxH.lineWidth = 1;
      ctxH.moveTo(x, 20);
      ctxH.lineTo(x, mm % 50 === 0 ? 8 : mm % 10 === 0 ? 12 : 16);
      ctxH.stroke();
      if (mm % 10 === 0) ctxH.fillText(mm, x + 2, 20);
    }
    rulerH.appendChild(canvasH);

    // Vertical
    const canvasV = document.createElement('canvas');
    canvasV.width = 20;
    canvasV.height = pageHeightPx + 80;
    canvasV.style.position = 'absolute';
    canvasV.style.left = '0';
    canvasV.style.top = '0';
    const ctxV = canvasV.getContext('2d');
    ctxV.fillStyle = '#606080';
    ctxV.font = '9px Arial';
    ctxV.textBaseline = 'top';
    const vStep = MM_TO_PX * this.zoom;
    for (let mm = 0; mm <= pageHeightMm; mm += 10) {
      const y = 40 + mm * vStep;
      ctxV.beginPath();
      ctxV.strokeStyle = '#606080';
      ctxV.lineWidth = 1;
      ctxV.moveTo(20, y);
      ctxV.lineTo(mm % 50 === 0 ? 8 : mm % 10 === 0 ? 12 : 16, y);
      ctxV.stroke();
      if (mm % 10 === 0) {
        ctxV.save();
        ctxV.translate(18, y + 2);
        ctxV.rotate(-Math.PI / 2);
        ctxV.fillText(mm, 0, 0);
        ctxV.restore();
      }
    }
    rulerV.appendChild(canvasV);
  }

  _syncRulersWithScroll() {
    if (!this.scrollArea) return;
    const hCanvas = document.querySelector('#ruler-h canvas');
    const vCanvas = document.querySelector('#ruler-v canvas');
    if (hCanvas) hCanvas.style.transform = `translateX(${-this.scrollArea.scrollLeft}px)`;
    if (vCanvas) vCanvas.style.transform = `translateY(${-this.scrollArea.scrollTop}px)`;
  }

  // ===== Render Elements =====
    renderElements() {
      // Remove all elements except the margin guide and grid-canvas
      this.pageCanvas.querySelectorAll('.canvas-element').forEach(el => el.remove());

    this.elements.forEach(el => {
      this._renderElement(el);
    });

      // Re-select if needed
      this._applySelectionToDOM();
    }

  _renderElement(el) {
    const domEl = document.createElement('div');
    domEl.className = 'canvas-element';
    domEl.dataset.id = el.id;
    domEl.dataset.type = el.type;

    const px = this._mmToPx(el.x);
    const py = this._mmToPx(el.y);
    const pw = this._mmToPx(el.width);
    const ph = this._mmToPx(el.height);

    domEl.style.left = px + 'px';
    domEl.style.top = py + 'px';
    domEl.style.width = pw + 'px';
    domEl.style.height = ph + 'px';

    const style = el.style || {};
    if (style.opacity !== undefined) domEl.style.opacity = style.opacity;

    switch (el.type) {
      case 'text':     this._buildTextEl(domEl, el); break;
        case 'field':    this._buildFieldEl(domEl, el); break;
        case 'user':     this._buildUserEl(domEl, el); break;
      case 'image':
      case 'logo':     this._buildImageEl(domEl, el); break;
        case 'rect':     this._buildRectEl(domEl, el); break;
      case 'line':     this._buildLineEl(domEl, el); break;
      case 'table':    this._buildTableEl(domEl, el); break;
      case 'datetime': this._buildDateTimeEl(domEl, el); break;
      case 'pagenum':  this._buildPageNumEl(domEl, el); break;
      case 'barcode':  this._buildBarcodeEl(domEl, el); break;
    }

    this.pageCanvas.appendChild(domEl);

    // Bind events
    domEl.addEventListener('mousedown', (e) => this._onElementMouseDown(e, el.id));
    domEl.addEventListener('dblclick', (e) => this._onElementDblClick(e, el.id));
    domEl.addEventListener('contextmenu', (e) => this._onElementContextMenu(e, el.id));

    return domEl;
  }

  _applyTextStyle(domEl, style) {
    if (!style) return;
    domEl.style.fontFamily = style.fontFamily || 'Arial';
    domEl.style.fontSize = (style.fontSize || 12) + 'pt';
    domEl.style.fontWeight = style.fontWeight || 'normal';
    domEl.style.fontStyle = style.fontStyle || 'normal';
    domEl.style.textDecoration = style.textDecoration || 'none';
    domEl.style.color = style.color || '#000000';
    domEl.style.textAlign = style.textAlign || 'left';
    domEl.style.backgroundColor = style.backgroundColor || 'transparent';
    domEl.style.overflow = 'hidden';
    domEl.style.wordBreak = 'break-word';
    // Use flex to honour textAlign visually on the element container
    domEl.style.display = 'flex';
    domEl.style.alignItems = 'flex-start';
    const ta = style.textAlign || 'left';
    if (ta === 'center') domEl.style.justifyContent = 'center';
    else if (ta === 'right') domEl.style.justifyContent = 'flex-end';
    else domEl.style.justifyContent = 'flex-start'; // left + justify
    if (style.borderWidth > 0) {
      domEl.style.border = `${style.borderWidth}px ${style.borderStyle || 'solid'} ${style.borderColor || '#000000'}`;
    } else {
      domEl.style.border = 'none';
    }
  }

  _buildTextEl(domEl, el) {
    domEl.classList.add('el-text');
    this._applyTextStyle(domEl, el.style);
    domEl.textContent = el.content || '';
    domEl.style.padding = '1px';
  }

  _buildFieldEl(domEl, el) {
    domEl.classList.add('el-field');
    this._applyTextStyle(domEl, el.style);
    const label = el.fieldName ? `{${el.fieldName}}` : '{Field}';
    domEl.textContent = label;
    domEl.style.padding = '1px';
    domEl.style.color = el.style?.color || '#3366cc';
  }

  _buildUserEl(domEl, el) {
    const inner = document.createElement('div');
    inner.style.width = '100%';
    inner.style.height = '100%';
    inner.style.display = 'flex';
    inner.style.alignItems = 'center';
    const align = el.style?.textAlign || 'left';
    inner.style.justifyContent = align === 'center' ? 'center' : (align === 'right' ? 'flex-end' : 'flex-start');
    const user = window.AuthStore?.currentUser?.();
    inner.textContent = user?.username || '{User}';
    domEl.appendChild(inner);
  }

  _buildImageEl(domEl, el) {
    domEl.classList.add('el-image');
    this._applyTextStyle(domEl, el.style);
    if (el.imageData) {
      const img = document.createElement('img');
      img.src = el.imageData;
      img.draggable = false;
      domEl.appendChild(img);
    } else {
      const ph = document.createElement('div');
      ph.className = 'el-image-placeholder';
      ph.textContent = '🖼';
      domEl.appendChild(ph);
    }
  }

  _buildRectEl(domEl, el) {
    domEl.classList.add('el-rect');
    const style = el.style || {};
    domEl.style.backgroundColor = style.backgroundColor || 'transparent';
    domEl.style.border = style.borderWidth > 0
      ? `${style.borderWidth}px ${style.borderStyle || 'solid'} ${style.borderColor || '#000000'}`
      : '1px solid #cccccc';
    domEl.style.opacity = style.opacity !== undefined ? style.opacity : 1;
  }

  _buildLineEl(domEl, el) {
    domEl.classList.add('el-line');
    const style = el.style || {};
    const color = style.borderColor || '#000000';
    const thickness = (style.borderWidth || 1);
    const direction = el.lineDirection || 'horizontal';
    const inner = document.createElement('div');
    inner.className = 'el-line-inner';
    inner.style.backgroundColor = color;
    if (direction === 'horizontal') {
      inner.style.width = '100%';
      inner.style.height = thickness + 'px';
    } else {
      inner.style.width = thickness + 'px';
      inner.style.height = '100%';
    }
    domEl.appendChild(inner);
  }

  _buildDateTimeEl(domEl, el) {
    domEl.classList.add('el-datetime');
    this._applyTextStyle(domEl, el.style);
    const fmt = el.datetimeFormat || 'DD/MM/YYYY';
    domEl.textContent = this._formatDateTime(fmt);
  }

  _buildPageNumEl(domEl, el) {
    domEl.classList.add('el-pagenum');
    this._applyTextStyle(domEl, el.style);
    const fmt = el.pagenumFormat || 'Page {n}';
    domEl.textContent = fmt.replace('{n}', '1').replace('{total}', '1');
  }

  _buildBarcodeEl(domEl, el) {
    domEl.classList.add('el-barcode');
    const bc = el.barcode || {};
    const value = el.fieldName
      ? `{${el.fieldName}}`
      : (el.content || '');

    if (typeof JsBarcode !== 'undefined') {
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      const sampleVal = (value && !value.startsWith('{')) ? value : '1234567890';
      try {
        const elHpx = this._mmToPx(el.height);
        const textAllowance = bc.showText !== false ? (bc.fontSize || 10) + 4 : 0;
        const barHeight = Math.max(8, Math.round(elHpx - textAllowance - 6));
        JsBarcode(svg, sampleVal, {
          format: bc.type || 'CODE128',
          displayValue: bc.showText !== false,
          fontSize: bc.fontSize || 10,
          lineColor: bc.textColor || '#000000',
          height: barHeight,
          width: 1.5,
          margin: 2,
          textMargin: 2,
        });
        // Fix SVG attrs so it fills the element bounds exactly
        const svgW = parseFloat(svg.getAttribute('width')) || 200;
        const svgH = parseFloat(svg.getAttribute('height')) || 60;
        if (!svg.getAttribute('viewBox')) svg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);
        svg.setAttribute('width', '100%');
        svg.setAttribute('height', '100%');
        svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
        svg.style.cssText = 'display:block;width:100%;height:100%;overflow:hidden;';
        domEl.style.overflow = 'hidden';
        domEl.appendChild(svg);
        if (value && value.startsWith('{')) {
          const lbl = document.createElement('div');
          lbl.style.cssText = `position:absolute;bottom:2px;right:4px;font-size:9px;color:var(--accent);font-style:italic;pointer-events:none;`;
          lbl.textContent = value;
          domEl.style.position = 'relative';
          domEl.appendChild(lbl);
        }
        return;
      } catch (e) { /* fall through to placeholder */ }
    }

    // Placeholder
    const ph = document.createElement('div');
    ph.className = 'el-barcode-placeholder';
    const bars = document.createElement('div');
    bars.className = 'el-barcode-placeholder-bars';
    [3,1,4,1,5,2,3,2,4,1,3,2,5,1,4,2,3,1].forEach(h => {
      const s = document.createElement('span');
      s.style.height = (h * 4) + 'px';
      bars.appendChild(s);
    });
    ph.appendChild(bars);
    const lbl = document.createElement('div');
    lbl.textContent = el.fieldName ? `{${el.fieldName}}` : (el.content || 'Set value');
    ph.appendChild(lbl);
    domEl.appendChild(ph);
  }

  _formatDateTime(format) {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const h12 = now.getHours() % 12 || 12;
    const tokens = {
      'YYYY': now.getFullYear(),
      'MMM':  months[now.getMonth()],
      'MM':   pad(now.getMonth() + 1),
      'DD':   pad(now.getDate()),
      'HH':   pad(now.getHours()),
      'hh':   pad(h12),
      'mm':   pad(now.getMinutes()),
      'ss':   pad(now.getSeconds()),
      'A':    now.getHours() < 12 ? 'AM' : 'PM',
    };
    return format.replace(/YYYY|MMM|MM|DD|HH|hh|mm|ss|A/g, m => tokens[m] !== undefined ? tokens[m] : m);
  }

  _buildTableEl(domEl, el) {
    domEl.classList.add('el-table');
    const tbl = el.table || { rows: 2, cols: 4, cells: [], theme: 'plain', borderMode: 'all', colWidths: [1, 1, 1, 1], detailMode: false };
    const rows = tbl.rows || 2;
    const cols = tbl.cols || 4;
    const footerEnabled = tbl.footerEnabled === true;
    const renderRows = rows + (footerEnabled ? 1 : 0);
    const theme = tbl.theme || 'plain';
    const borderMode = tbl.borderMode || 'all';

    // colWidths support
    const colWidths = tbl.colWidths && tbl.colWidths.length === cols ? tbl.colWidths : Array(cols).fill(1);
    const total = colWidths.reduce((s, w) => s + w, 0);

    // Theme presets
    const themes = {
      'plain':       { headerBg: '#ffffff', headerColor: '#000', rowBg: '#fff', altRowBg: '#ffffff' },
      'dark-header': { headerBg: '#2d2d4e', headerColor: '#fff', rowBg: '#fff', altRowBg: '#f0f0f8' },
      'blue':        { headerBg: '#2a5298', headerColor: '#fff', rowBg: '#fff', altRowBg: '#eef3fb' },
      'green':       { headerBg: '#2d6a4f', headerColor: '#fff', rowBg: '#fff', altRowBg: '#eef7f2' },
      'striped':     { headerBg: '#555', headerColor: '#fff', rowBg: '#fff', altRowBg: '#f5f5f5' },
    };

    const t = themes[theme] || themes['plain'];
    const headerBg = tbl.headerBg || t.headerBg;
    const headerColor = tbl.headerColor || t.headerColor;
    const rowBg = tbl.rowBg || t.rowBg;
    const altRowBg = tbl.altRowBg || t.altRowBg;

    // rowHeights support
    const rowHeights = this._baseRowHeightsMm(tbl, el.height);
    const footerRowHeight = Number.isFinite(tbl.footerRowHeight) && tbl.footerRowHeight > 0
      ? tbl.footerRowHeight
      : (rowHeights[Math.max(1, rows) - 1] || 1);
    const renderRowHeights = footerEnabled ? rowHeights.concat([footerRowHeight]) : rowHeights;
    const rowTotal = renderRowHeights.reduce((s, h) => s + h, 0);

    domEl.style.display = 'grid';
    domEl.style.gridTemplateColumns = colWidths.map(w => (w / total * 100).toFixed(2) + '%').join(' ');
    domEl.style.gridTemplateRows = renderRowHeights.map(h => (h / rowTotal * 100).toFixed(2) + '%').join(' ');
    domEl.style.overflow = 'hidden';
    domEl.style.position = 'relative';

    const style = el.style || {};
    const bc = style.borderColor || '#cccccc';
    const bw = Math.max(1, style.borderWidth !== undefined ? style.borderWidth : 1);

    for (let r = 0; r < renderRows; r++) {
      for (let c = 0; c < cols; c++) {
        const isFooterRow = footerEnabled && r === renderRows - 1;
        const cellData = !isFooterRow ? (tbl.cells || []).find(cl => cl.row === r && cl.col === c) : null;
        const cell = document.createElement('div');
        cell.className = 'el-table-cell';
        cell.dataset.row = r;
        cell.dataset.col = c;
        cell.style.fontFamily = style.fontFamily || 'Arial';
        cell.style.fontSize = (style.fontSize || 10) + 'pt';
        cell.style.overflow = 'hidden';
        cell.style.padding = '2px 4px';
        cell.style.boxSizing = 'border-box';
        cell.style.display = 'flex';
        cell.style.alignItems = 'center';

        // Apply column props (alignment + padding)
        const colProps = tbl.colProps || [];
        const cp = colProps[c] || {};
        const colAlign = cp.textAlign || style.textAlign || 'left';
        cell.style.justifyContent = colAlign === 'center' ? 'center' : (colAlign === 'right' ? 'flex-end' : 'flex-start');
        cell.style.textAlign = colAlign;
        if (cp.paddingLeft !== undefined) cell.style.paddingLeft = cp.paddingLeft + 'px';
        if (cp.paddingRight !== undefined) cell.style.paddingRight = cp.paddingRight + 'px';

        const cellStyle = cellData?.style || {};
        if (cellStyle.fontFamily) cell.style.fontFamily = cellStyle.fontFamily;
        if (cellStyle.fontSize) cell.style.fontSize = cellStyle.fontSize + 'pt';
        if (cellStyle.fontStyle) cell.style.fontStyle = cellStyle.fontStyle;
        if (cellStyle.textDecoration) cell.style.textDecoration = cellStyle.textDecoration;

        // Row background
        if (r === 0) {
          cell.style.backgroundColor = headerBg;
          cell.style.color = cellStyle.color || headerColor;
          cell.style.fontWeight = cellStyle.fontWeight || '600';
        } else if (isFooterRow) {
          cell.style.backgroundColor = rowBg;
          cell.style.color = cellStyle.color || style.color || '#000000';
          cell.style.fontWeight = cellStyle.fontWeight || '600';
        } else {
          cell.style.backgroundColor = (r % 2 === 0) ? altRowBg : rowBg;
          cell.style.color = cellStyle.color || style.color || '#000000';
          cell.style.fontWeight = cellStyle.fontWeight || style.fontWeight || 'normal';
        }

        // Cell borders
        if (borderMode === 'all') {
          cell.style.border = `${bw}px ${style.borderStyle || 'solid'} ${bc}`;
        } else if (borderMode === 'outer') {
          cell.style.border = 'none';
        } else if (borderMode === 'header-outer') {
          cell.style.border = 'none';
          if (r === 0) cell.style.borderBottom = `2px ${style.borderStyle || 'solid'} ${bc}`;
          if (isFooterRow) cell.style.borderTop = `2px ${style.borderStyle || 'solid'} ${bc}`;
        } else {
          cell.style.border = 'none';
        }

        // Content — barcode columns show a barcode preview in data rows
        if (cp.barcode && r > 0 && !isFooterRow) {
          cell.style.padding = '2px';
          cell.style.alignItems = 'center';
          cell.style.justifyContent = 'center';
          const sampleVal = (cellData?.fieldName) ? '1234567890' : (cellData?.content || '1234567890');
          if (typeof JsBarcode !== 'undefined') {
            const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
            try {
              JsBarcode(svg, sampleVal, {
                format: cp.barcodeType || 'CODE128',
                displayValue: cp.barcodeShowText !== false,
                height: 18, width: 1, margin: 1, fontSize: 7, textMargin: 1,
              });
              const svgW = parseFloat(svg.getAttribute('width')) || 100;
              const svgH = parseFloat(svg.getAttribute('height')) || 28;
              if (!svg.getAttribute('viewBox')) svg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);
              svg.setAttribute('width', '100%');
              svg.setAttribute('height', '100%');
              svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
              svg.style.cssText = 'display:block;width:100%;height:100%;';
              cell.appendChild(svg);
            } catch(e) {
              cell.textContent = sampleVal;
            }
          } else {
            cell.textContent = sampleVal;
          }
          // Show field name as small overlay label
          if (cellData?.fieldName) {
            const lbl = document.createElement('span');
            lbl.style.cssText = 'position:absolute;bottom:1px;right:2px;font-size:8px;color:#3366cc;pointer-events:none;';
            lbl.textContent = `{${cellData.fieldName}}`;
            cell.style.position = 'relative';
            cell.appendChild(lbl);
          }
        } else if (isFooterRow) {
          const footerType = cp.footerType || 'none';
          if (footerType === 'text') {
            cell.textContent = cp.footerText || '';
          } else if (footerType === 'sum') {
            cell.textContent = 'SUM';
          } else {
            cell.textContent = '';
          }
        } else if (cellData && cellData.fieldName) {
          cell.textContent = `{${cellData.fieldName}}`;
          cell.style.color = r === 0 ? headerColor : (style.color || '#000000');
          cell.classList.add('has-field');
        } else if (cellData && cellData.content) {
          cell.textContent = cellData.content;
        } else {
          cell.textContent = '';
        }

        if (r === 0) {
          const insertBtn = document.createElement('button');
          insertBtn.type = 'button';
          insertBtn.className = 'table-col-insert-btn';
          insertBtn.title = 'Insert column after';
          insertBtn.textContent = '+';
          insertBtn.addEventListener('click', (ev) => {
            ev.preventDefault();
            ev.stopPropagation();
            this.selectElement(el.id);
            this.insertTableColumn(el.id, c, 'after');
          });
          cell.style.position = 'relative';
          cell.appendChild(insertBtn);
        }

        // Single-click in header selects the whole column (for barcode/column props).
        // Use Ctrl+click on header when you want only that one header cell selected.
        cell.addEventListener('click', (e) => {
          e.stopPropagation();
          if (this.selectedId === el.id) {
            if ((r === 0 || isFooterRow) && !e.ctrlKey) {
              this._selectTableColumn(el.id, c);
            } else {
              this._selectTableCell(el.id, r, c);
            }
          }
        });
        cell.addEventListener('contextmenu', () => {
          if (this.selectedId === el.id && r === 0) {
            this._selectTableColumn(el.id, c);
          }
        });

        // Make cell a drop target for fields
        if (!isFooterRow) {
          cell.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.stopPropagation();
            cell.classList.add('drop-target');
          });
          cell.addEventListener('dragleave', () => cell.classList.remove('drop-target'));
          cell.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            cell.classList.remove('drop-target');
            const itemType = e.dataTransfer.getData('application/x-printmore-item-type') || 'field';
            const itemName = e.dataTransfer.getData('text/plain') || this.fieldDragName;
            if (itemName) {
              if (itemType === 'text') this._updateTableCell(el.id, r, c, null, itemName);
              else this._updateTableCell(el.id, r, c, itemName, null);
            }
          });
        }

        // Double-click to edit cell text as heading
        cell.addEventListener('dblclick', (e) => {
          if (isFooterRow) return;
          e.stopPropagation();
          const editor = document.createElement('input');
          editor.type = 'text';
          editor.value = cellData?.content || (cellData?.fieldName ? `{${cellData.fieldName}}` : '');
          editor.style.cssText = `width:100%;height:100%;border:none;background:rgba(255,255,255,0.9);color:#000;font:inherit;padding:0 2px;box-sizing:border-box;`;
          cell.innerHTML = '';
          cell.appendChild(editor);
          editor.focus();
          editor.select();
          const finishCellEdit = () => {
            const val = editor.value.trim();
            const fieldRef = val.match(/^\{(.+)\}$/);
            if (fieldRef) {
              this._updateTableCell(el.id, r, c, fieldRef[1], null);
            } else {
              this._updateTableCell(el.id, r, c, null, val);
            }
          };
          editor.addEventListener('blur', finishCellEdit);
          editor.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') { ev.preventDefault(); editor.blur(); }
            if (ev.key === 'Escape') { editor.value = ''; editor.blur(); }
          });
        });

        domEl.appendChild(cell);
      }
    }

    // Add column and row resize handles after all cells
    this._addColResizeHandles(domEl, el);
    this._addRowResizeHandles(domEl, el);
  }

  // ===== Column Resize =====
  _addColResizeHandles(tableDiv, el) {
    const tbl = el.table;
    const cols = tbl.cols || 3;
    if (cols < 1) return;
    const colWidthsMm = this._colWidthsMm(tbl, el.width);
    const total = colWidthsMm.reduce((s, w) => s + w, 0);

    // Handle at right edge of EVERY column (including last).
    // Dragging any handle resizes only that column; table total width adjusts.
    // The last column's handle sits at calc(100% - 3px) — just inside the table border.
    let cumPct = 0;
    for (let c = 0; c < cols; c++) {
      cumPct += (colWidthsMm[c] / total) * 100;
      const isLast = c === cols - 1;
      const leftCSS = isLast ? 'calc(100% - 3px)' : `calc(${cumPct}% - 3px)`;

      const handle = document.createElement('div');
      handle.className = 'col-resize-handle';
      handle.dataset.colIndex = c;
      handle.dataset.elementId = el.id;
      handle.style.cssText = `position:absolute;top:0;left:${leftCSS};width:6px;height:100%;cursor:col-resize;z-index:150;background:transparent;`;

      handle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        e.preventDefault();
        this._startColResize(e, el.id, c);
      });
      handle.addEventListener('mouseenter', () => handle.style.background = 'rgba(108,99,255,0.35)');
      handle.addEventListener('mouseleave', () => handle.style.background = 'transparent');

      tableDiv.appendChild(handle);
    }
  }

  // Convert colWidths (may be proportional or absolute mm) to absolute mm values.
  _colWidthsMm(tbl, elWidthMm) {
    const cols = tbl.cols || 3;
    const stored = tbl.colWidths;
    if (!stored || stored.length !== cols) return Array(cols).fill(elWidthMm / cols);
    const sum = stored.reduce((s, w) => s + w, 0);
    // If sum < 20 assume proportional (old format), convert to mm
    return sum < 20 ? stored.map(w => (w / sum) * elWidthMm) : stored.slice();
  }

  _startColResize(e, elementId, colIndex, startWidths) {
    const el = this._findElement(elementId);
    this.colResizeState = {
      elementId,
      colIndex,
      startX: e.clientX,
      // Store as absolute mm for unambiguous arithmetic
      startWidthsMm: this._colWidthsMm(el.table, el.width),
      startElWidth: el.width,
    };
  }

  _performColResize(e) {
    const rs = this.colResizeState;
    if (!rs) return;
    const el = this._findElement(rs.elementId);
    if (!el) return;

    const dxMm = this._pxToMm(e.clientX - rs.startX);
    const minColMm = 5; // 5 mm minimum column width

    // Only resize the dragged column; table width grows/shrinks to compensate.
    const newMms = rs.startWidthsMm.slice();
    newMms[rs.colIndex] = Math.max(minColMm, newMms[rs.colIndex] + dxMm);
    const actualDxMm = newMms[rs.colIndex] - rs.startWidthsMm[rs.colIndex];

    el.table.colWidths = newMms;
    el.width = Math.max(20, rs.startElWidth + actualDxMm);

    // Update DOM
    const domEl = this.pageCanvas.querySelector(`[data-id="${rs.elementId}"]`);
    if (domEl) {
      domEl.style.width = this._mmToPx(el.width) + 'px';
      const total = newMms.reduce((s, w) => s + w, 0);
      domEl.style.gridTemplateColumns = newMms.map(w => (w / total * 100).toFixed(2) + '%').join(' ');
      // Update handle positions (one handle per inter-column gap)
      let cumPct = 0;
      domEl.querySelectorAll('.col-resize-handle').forEach((h, hi) => {
        cumPct += (newMms[hi] / total) * 100;
        h.style.left = `calc(${cumPct}% - 3px)`;
      });
    }
  }

  // ===== Row Resize =====
  _addRowResizeHandles(tableDiv, el) {
    const tbl = el.table;
    const rowHeightsMm = this._rowHeightsMm(tbl, el.height);
    const rows = rowHeightsMm.length;
    if (rows < 1) return;
    const total = rowHeightsMm.reduce((s, h) => s + h, 0);

    // Handle at bottom edge of EVERY row (including last = table bottom border).
    let cumPct = 0;
    for (let r = 0; r < rows; r++) {
      cumPct += (rowHeightsMm[r] / total) * 100;
      const isLast = r === rows - 1;
      const topCSS = isLast ? 'calc(100% - 3px)' : `calc(${cumPct}% - 3px)`;

      const handle = document.createElement('div');
      handle.className = 'row-resize-handle';
      handle.dataset.rowIndex = r;
      handle.style.cssText = `position:absolute;left:0;top:${topCSS};height:6px;width:100%;cursor:row-resize;z-index:150;background:transparent;`;
      handle.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        e.preventDefault();
        this._startRowResize(e, el.id, r);
      });
      handle.addEventListener('mouseenter', () => handle.style.background = 'rgba(108,99,255,0.35)');
      handle.addEventListener('mouseleave', () => handle.style.background = 'transparent');
      tableDiv.appendChild(handle);
    }
  }

  // Convert rowHeights (may be proportional or absolute mm) to absolute mm values.
  _baseRowHeightsMm(tbl, elHeightMm) {
    const rows = tbl.rows || 3;
    const stored = tbl.rowHeights;
    if (!stored || stored.length !== rows) return Array(rows).fill(elHeightMm / rows);
    const sum = stored.reduce((s, h) => s + h, 0);
    return sum < 20 ? stored.map(h => (h / sum) * elHeightMm) : stored.slice();
  }

  // Render row heights include footer row (if enabled).
  _rowHeightsMm(tbl, elHeightMm) {
    const base = this._baseRowHeightsMm(tbl, elHeightMm);
    if (!tbl.footerEnabled) return base;
    const fallback = base[Math.max(1, base.length) - 1] || (elHeightMm / Math.max(2, (tbl.rows || 1) + 1));
    const footer = Number.isFinite(tbl.footerRowHeight) && tbl.footerRowHeight > 0 ? tbl.footerRowHeight : fallback;
    return base.concat([footer]);
  }

  _startRowResize(e, elementId, rowIndex) {
    const el = this._findElement(elementId);
    this.rowResizeState = {
      elementId, rowIndex, startY: e.clientY,
      startHeightsMm: this._rowHeightsMm(el.table, el.height),
      startElHeight: el.height,
    };
  }

  _performRowResize(e) {
    const rs = this.rowResizeState;
    if (!rs) return;
    const el = this._findElement(rs.elementId);
    if (!el) return;

    const dyMm = this._pxToMm(e.clientY - rs.startY);
    const minRowMm = 3; // 3 mm minimum row height

    // Only resize the dragged row; table height grows/shrinks to compensate.
    const newMms = rs.startHeightsMm.slice();
    newMms[rs.rowIndex] = Math.max(minRowMm, newMms[rs.rowIndex] + dyMm);
    const actualDyMm = newMms[rs.rowIndex] - rs.startHeightsMm[rs.rowIndex];

    const baseRows = el.table.rows || 3;
    if (el.table.footerEnabled) {
      el.table.rowHeights = newMms.slice(0, baseRows);
      el.table.footerRowHeight = newMms[baseRows] || el.table.footerRowHeight || newMms[Math.max(0, baseRows - 1)] || 1;
    } else {
      el.table.rowHeights = newMms;
    }
    el.height = Math.max(10, rs.startElHeight + actualDyMm);

    const domEl = this.pageCanvas.querySelector(`[data-id="${rs.elementId}"]`);
    if (domEl) {
      domEl.style.height = this._mmToPx(el.height) + 'px';
      const total = newMms.reduce((s, h) => s + h, 0);
      domEl.style.gridTemplateRows = newMms.map(h => (h / total * 100).toFixed(2) + '%').join(' ');
      let cumPct = 0;
      domEl.querySelectorAll('.row-resize-handle').forEach((h, ri) => {
        cumPct += (newMms[ri] / total) * 100;
        const isLast = ri === newMms.length - 1;
        h.style.top = isLast ? 'calc(100% - 3px)' : `calc(${cumPct}% - 3px)`;
      });
    }
  }

  // ===== Table Cell Selection =====
  _selectTableCell(elementId, row, col) {
    this.selectedCell = { elementId, row, col };
    const el = this._findElement(elementId);
    const cellData = el?.table?.cells?.find(c => c.row === row && c.col === col);
    const cellStyle = cellData?.style || {};
    if (cellStyle.fontFamily) document.getElementById('prop-font-family').value = cellStyle.fontFamily;
    if (cellStyle.fontSize) document.getElementById('prop-font-size').value = cellStyle.fontSize;
    if (cellStyle.color) document.getElementById('prop-color').value = cellStyle.color;
    this._setStyleBtnActive('prop-bold', cellStyle.fontWeight === 'bold');
    this._setStyleBtnActive('prop-italic', cellStyle.fontStyle === 'italic');
    this._setStyleBtnActive('prop-underline', cellStyle.textDecoration === 'underline');

    // Hide column properties panel when a non-header cell is clicked
    document.getElementById('prop-group-column')?.classList.add('hidden');

    // Highlight selected cell
    const domEl = this.pageCanvas.querySelector(`[data-id="${elementId}"]`);
    if (domEl) {
      domEl.querySelectorAll('.el-table-cell').forEach(c => c.style.outline = 'none');
      const cellDom = domEl.querySelector(`.el-table-cell[data-row="${row}"][data-col="${col}"]`);
      if (cellDom) cellDom.style.outline = '2px solid var(--accent)';
    }
  }

  _selectTableColumn(elementId, colIndex) {
    this.selectedCell = { elementId, row: 0, col: colIndex, isColumn: true };

    // Highlight entire column
    const domEl = this.pageCanvas.querySelector(`[data-id="${elementId}"]`);
    if (domEl) {
      domEl.querySelectorAll('.el-table-cell').forEach(c => c.style.outline = 'none');
      domEl.querySelectorAll(`.el-table-cell[data-col="${colIndex}"]`).forEach(c => {
        c.style.outline = '2px solid var(--accent)';
      });
    }

    // Show column properties panel
    this._showColumnProperties(elementId, colIndex);
  }

  _showColumnProperties(elementId, colIndex) {
    // Show column panel (keep other groups as-is)
    document.getElementById('prop-group-column')?.classList.remove('hidden');

    const el = this._findElement(elementId);
    if (!el) return;

    // Update label
    const label = document.getElementById('col-prop-label');
    if (label) label.textContent = `Column ${colIndex + 1}`;

    // Load existing colProps
    const colProps = el.table.colProps || [];
    const cp = colProps[colIndex] || {};
    const align = cp.textAlign || 'left';
    const padLeft = cp.paddingLeft !== undefined ? cp.paddingLeft : 5;
    const padRight = cp.paddingRight !== undefined ? cp.paddingRight : 5;

    // Update align buttons
    document.querySelectorAll('#col-align-btns .align-btn[data-col-align]').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.colAlign === align);
    });

    // Update padding inputs
    const padLEl = document.getElementById('col-prop-pad-left');
    const padREl = document.getElementById('col-prop-pad-right');
    if (padLEl) padLEl.value = padLeft;
    if (padREl) padREl.value = padRight;

    // Barcode column toggle
    const barcodeChk = document.getElementById('col-prop-barcode');
    const barcodeOpts = document.getElementById('col-barcode-opts');
    if (barcodeChk) {
      barcodeChk.checked = !!cp.barcode;
      if (barcodeOpts) barcodeOpts.classList.toggle('hidden', !cp.barcode);
    }
    const barcodeTypeEl = document.getElementById('col-prop-barcode-type');
    if (barcodeTypeEl) barcodeTypeEl.value = cp.barcodeType || 'CODE128';
    const barcodeShowTextEl = document.getElementById('col-prop-barcode-showtext');
    if (barcodeShowTextEl) barcodeShowTextEl.checked = cp.barcodeShowText !== false;
    const footerTypeEl = document.getElementById('col-prop-footer-type');
    if (footerTypeEl) footerTypeEl.value = cp.footerType || 'none';
    const footerTextEl = document.getElementById('col-prop-footer-text');
    if (footerTextEl) footerTextEl.value = cp.footerText || '';
    const footerMergeNextEl = document.getElementById('col-prop-footer-merge-next');
    if (footerMergeNextEl) footerMergeNextEl.checked = !!cp.footerMergeNext;
  }

  // ===== Resize Handles =====
  _renderResizeHandles(domEl, elementId) {
    domEl.querySelectorAll('.resize-handle').forEach(h => h.remove());
    const handles = ['nw','n','ne','w','e','sw','s','se'];
    handles.forEach(pos => {
      const h = document.createElement('div');
      h.className = `resize-handle ${pos}`;
      h.dataset.handle = pos;
      h.addEventListener('mousedown', (e) => {
        e.stopPropagation();
        this.startResize(e, elementId, pos);
      });
      domEl.appendChild(h);
    });
  }

  _removeResizeHandles() {
    this.pageCanvas.querySelectorAll('.resize-handle').forEach(h => h.remove());
  }

  // ===== Tools Panel =====
  initToolbarEvents() {
    document.querySelectorAll('.tool-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        this.setActiveTool(btn.dataset.tool);
      });
    });

    document.querySelectorAll('.side-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        const target = tab.dataset.sideTab;
        document.querySelectorAll('.side-tab').forEach(t => t.classList.toggle('active', t === tab));
        document.getElementById('fields-list')?.classList.toggle('hidden', target !== 'fields');
        document.getElementById('texts-list')?.classList.toggle('hidden', target !== 'texts');
        document.querySelector('.field-list-controls:not(#text-list-controls)')?.classList.toggle('hidden', target !== 'fields');
        document.getElementById('text-list-controls')?.classList.toggle('hidden', target !== 'texts');
      });
    });

    document.getElementById('btn-add-field')?.addEventListener('click', () => {
      const input = document.getElementById('field-add-input');
      const value = input.value.trim();
      if (!value) return;
      this.layout.fields = this.layout.fields || [];
      if (!this.layout.fields.some(f => f.toLowerCase() === value.toLowerCase())) {
        this.layout.fields.push(value);
        this._saveLayoutDefinition();
      }
      input.value = '';
    });

    document.getElementById('btn-add-text')?.addEventListener('click', () => {
      const input = document.getElementById('text-add-input');
      const value = input.value.trim();
      if (!value) return;
      this.layout.texts = this.layout.texts || [];
      if (!this.layout.texts.some(t => t.toLowerCase() === value.toLowerCase())) {
        this.layout.texts.push(value);
        this._saveLayoutDefinition();
      }
      input.value = '';
    });

    this._updateToolButtons();
    this._initPropertiesEvents();
    this._initContextMenu();

    // Inline page settings (shown when no element is selected)
    const inlinePageChange = () => {
      if (!this.layout) return;
      const sz = document.getElementById('inline-page-size')?.value;
      const or = document.getElementById('inline-orientation')?.value;
      const mt = parseFloat(document.getElementById('inline-margin-top')?.value) || 0;
      const mr = parseFloat(document.getElementById('inline-margin-right')?.value) || 0;
      const mb = parseFloat(document.getElementById('inline-margin-bottom')?.value) || 0;
      const ml = parseFloat(document.getElementById('inline-margin-left')?.value) || 0;
      if (sz) this.layout.page.size = sz;
      if (or) this.layout.page.orientation = or;
      if ((sz || this.layout.page.size) === 'custom') {
        this.layout.page.customWidthMm = Math.max(20, parseFloat(document.getElementById('inline-custom-width')?.value) || 210);
        this.layout.page.customHeightMm = Math.max(20, parseFloat(document.getElementById('inline-custom-height')?.value) || 297);
      }
      this.layout.page.marginTop = mt;
      this.layout.page.marginRight = mr;
      this.layout.page.marginBottom = mb;
      this.layout.page.marginLeft = ml;
      // Zone settings
      this.layout.page.headerEnabled = document.getElementById('inline-header-enabled')?.checked ?? false;
      this.layout.page.headerHeight  = parseFloat(document.getElementById('inline-header-height')?.value) || 20;
      this.layout.page.headerPages   = document.getElementById('inline-header-pages')?.value || 'all';
      this.layout.page.footerEnabled = document.getElementById('inline-footer-enabled')?.checked ?? false;
      this.layout.page.footerHeight  = parseFloat(document.getElementById('inline-footer-height')?.value) || 15;
      this.layout.page.footerPages   = document.getElementById('inline-footer-pages')?.value || 'all';
      this.initCanvas();
      this.renderElements();
      this._showNoSelectionPanel();
      this.saveLayout();
    };
    ['inline-page-size','inline-custom-width','inline-custom-height','inline-orientation','inline-margin-top','inline-margin-right','inline-margin-bottom','inline-margin-left',
     'inline-header-height','inline-header-pages','inline-footer-height','inline-footer-pages'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', inlinePageChange);
      if (el) el.addEventListener('input', inlinePageChange);
    });
    document.getElementById('inline-page-size')?.addEventListener('change', () => {
      const isCustom = document.getElementById('inline-page-size')?.value === 'custom';
      document.getElementById('inline-custom-size-group')?.classList.toggle('hidden', !isCustom);
    });

    // Zone enable/disable toggles
    document.getElementById('inline-header-enabled')?.addEventListener('change', (e) => {
      document.getElementById('inline-header-options')?.classList.toggle('hidden', !e.target.checked);
      inlinePageChange();
    });
    document.getElementById('inline-footer-enabled')?.addEventListener('change', (e) => {
      document.getElementById('inline-footer-options')?.classList.toggle('hidden', !e.target.checked);
      inlinePageChange();
    });

  }

  setActiveTool(tool) {
    this.activeTool = tool;
    this._updateToolButtons();
    this.pageCanvas.style.cursor = tool === 'select' ? 'default' : 'crosshair';
    // Deselect when switching to non-select tool
    if (tool !== 'select') {
      this.deselectAll();
    }
  }

  _updateToolButtons() {
    document.querySelectorAll('.tool-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.tool === this.activeTool);
    });
  }

  // ===== Fields List =====
  renderFieldsList() {
    const renderList = (listId, items, type) => {
      const list = document.getElementById(listId);
      if (!list) return;
      list.innerHTML = '';
      if (items.length === 0) {
        list.innerHTML = `<div style="color:var(--text-muted);font-size:11px;padding:4px;">No ${type === 'field' ? 'fields' : 'text'} defined.</div>`;
        return;
      }
      items.forEach((name, index) => {
      const item = document.createElement('div');
      item.className = 'field-item';
      item.draggable = true;
        item.dataset.name = name;
        item.dataset.type = type;
        const asTextButton = type === 'field' ? '<button data-action="as-text" title="Drag as text">T</button>' : '';
        item.innerHTML = `
          <span class="field-item-icon">${type === 'field' ? '{}' : 'T'}</span>
          <span class="field-item-name">${window.escapeHtml(name)}</span>
          <span class="field-item-actions">
            ${asTextButton}
            <button data-action="up" title="Move up">↑</button>
            <button data-action="down" title="Move down">↓</button>
          </span>
        `;

      item.addEventListener('dragstart', (e) => {
          this.fieldDragName = name;
        e.dataTransfer.effectAllowed = 'copy';
          e.dataTransfer.setData('text/plain', name);
          e.dataTransfer.setData('application/x-printmore-item-type', type);
      });
      item.addEventListener('dragend', () => {
        this.fieldDragName = null;
      });
        const asTextBtn = type === 'field' ? item.querySelector('[data-action="as-text"]') : null;
        if (asTextBtn) {
          asTextBtn.setAttribute('draggable', 'true');
          asTextBtn.addEventListener('mousedown', (e) => e.stopPropagation());
          asTextBtn.addEventListener('dragstart', (e) => {
            this.fieldDragName = name;
            e.dataTransfer.effectAllowed = 'copy';
            e.dataTransfer.setData('text/plain', name);
            e.dataTransfer.setData('application/x-printmore-item-type', 'text');
          });
        }
        item.querySelector('[data-action="up"]').addEventListener('click', (e) => {
          e.stopPropagation();
          if (index <= 0) return;
          [items[index - 1], items[index]] = [items[index], items[index - 1]];
          this._saveLayoutDefinition();
        });
        item.querySelector('[data-action="down"]').addEventListener('click', (e) => {
          e.stopPropagation();
          if (index >= items.length - 1) return;
          [items[index + 1], items[index]] = [items[index], items[index + 1]];
          this._saveLayoutDefinition();
        });

      list.appendChild(item);
    });
    };

    const textItems = (this.layout.texts && this.layout.texts.length)
      ? this.layout.texts
      : (this.layout.fields || []);
    renderList('fields-list', this.layout.fields || [], 'field');
    renderList('texts-list', textItems, 'text');
  }

  _saveLayoutDefinition() {
    this.layout.defaultStyle = { ...this.defaultStyle };
    this.layout.elements = JSON.parse(JSON.stringify(this.elements));
    window.saveLayout(this.layout);
    this.renderFieldsList();
    this._populateFieldSelects();
  }

  // ===== Canvas Events =====
    initCanvasEvents() {
      this.pageCanvas.addEventListener('mousedown', (e) => this.handleCanvasMouseDown(e));
      this.pageCanvas.addEventListener('contextmenu', (e) => {
        if (this.suppressNextContextMenu) {
          e.preventDefault();
          e.stopPropagation();
          this.suppressNextContextMenu = false;
        }
      });
      this.pageCanvas.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
    this.pageCanvas.addEventListener('drop', (e) => this._onCanvasDrop(e));

    // Hide context menu on outside click
    document.addEventListener('click', () => this._hideContextMenu());
  }

  _onCanvasDrop(e) {
    e.preventDefault();
    const itemType = e.dataTransfer.getData('application/x-printmore-item-type') || 'field';
    const itemName = e.dataTransfer.getData('text/plain') || this.fieldDragName;
    if (!itemName) return;

    const rect = this.pageCanvas.getBoundingClientRect();
    const xPx = e.clientX - rect.left;
    const yPx = e.clientY - rect.top;
    const xMm = this._pxToMm(xPx);
    const yMm = this._pxToMm(yPx);

    this.saveToHistory();
    const el = this._createElementData(itemType === 'text' ? 'text' : 'field', xMm, yMm);
    if (itemType === 'text') {
      el.content = itemName;
      this._autoFitTextElement(el);
    } else {
      el.fieldName = itemName;
      this._autoFitTextElement(el);
    }
    this.elements.push(el);
    this.renderElements();
    this.selectElement(el.id);
    this._saveLayoutElements();
  }

  handleCanvasMouseDown(e) {
    if (this.activeTool === 'select' && e.button === 2) {
      this._startMarqueeSelect(e);
      return;
    }
    if (e.button !== 0) return;

    // Click on empty canvas
    const target = e.target;
    const isCanvasElement = target.closest('.canvas-element');
    if (isCanvasElement) return; // handled by element event

    if (this.activeTool === 'select') {
      this._startMarqueeSelect(e);
      return;
    }

    // Deselect
    this.deselectAll();

    // Start drawing a new element
    const rect = this.pageCanvas.getBoundingClientRect();
    const xPx = e.clientX - rect.left;
    const yPx = e.clientY - rect.top;

    this.drawState = {
      startX: xPx,
      startY: yPx,
      tool: this.activeTool,
      ghostEl: null,
    };

    // Create ghost
    const ghost = document.createElement('div');
    ghost.style.cssText = `position:absolute;border:2px dashed var(--accent);background:rgba(108,99,255,0.08);pointer-events:none;z-index:999;box-sizing:border-box;`;
    ghost.style.left = xPx + 'px';
    ghost.style.top = yPx + 'px';
    ghost.style.width = '0px';
    ghost.style.height = '0px';
    this.pageCanvas.appendChild(ghost);
    this.drawState.ghostEl = ghost;

    e.preventDefault();
  }

  handleCanvasMouseMove(e) {
    if (this.colResizeState) { this._performColResize(e); return; }
    if (this.rowResizeState) { this._performRowResize(e); return; }
    if (this.marqueeState) {
      this._performMarqueeSelect(e);
    } else if (this.dragState) {
      this._performDrag(e);
    } else if (this.resizeState) {
      this._performResize(e);
    } else if (this.drawState) {
      this._performDraw(e);
    }
  }

  handleCanvasMouseUp(e) {
    if (this.colResizeState) {
      this.saveToHistory(); this._saveLayoutElements(); this.colResizeState = null; return;
    }
    if (this.rowResizeState) {
      this.saveToHistory(); this._saveLayoutElements(); this.rowResizeState = null; return;
    }
    if (this.marqueeState) { this._finishMarqueeSelect(e); return; }
    if (this.drawState) { this._finishDraw(e); }
    if (this.dragState) { this._finishDrag(); }
    if (this.resizeState) { this._finishResize(); }
  }

  _onMouseMove(e) { this.handleCanvasMouseMove(e); }
  _onMouseUp(e) { this.handleCanvasMouseUp(e); }

  _performDraw(e) {
    if (!this.drawState || !this.drawState.ghostEl) return;
    const rect = this.pageCanvas.getBoundingClientRect();
    const xPx = e.clientX - rect.left;
    const yPx = e.clientY - rect.top;
    const x = Math.min(xPx, this.drawState.startX);
    const y = Math.min(yPx, this.drawState.startY);
    const w = Math.abs(xPx - this.drawState.startX);
    const h = Math.abs(yPx - this.drawState.startY);
    const ghost = this.drawState.ghostEl;
    ghost.style.left = x + 'px';
    ghost.style.top = y + 'px';
    ghost.style.width = w + 'px';
    ghost.style.height = h + 'px';
  }

    _finishDraw(e) {
    if (!this.drawState) return;
    const ghost = this.drawState.ghostEl;
    if (ghost) ghost.remove();

    const rect = this.pageCanvas.getBoundingClientRect();
    const xPx = e.clientX - rect.left;
    const yPx = e.clientY - rect.top;
    const sx = Math.min(xPx, this.drawState.startX);
    const sy = Math.min(yPx, this.drawState.startY);
    const w = Math.abs(xPx - this.drawState.startX);
    const h = Math.abs(yPx - this.drawState.startY);

    const xMm = this._pxToMm(sx);
    const yMm = this._pxToMm(sy);
    let wMm = this._pxToMm(w);
    let hMm = this._pxToMm(h);

    // Minimum size
    if (wMm < MIN_ELEMENT_SIZE) {
      if (this.drawState?.tool === 'text' || this.drawState?.tool === 'field') {
        wMm = this._defaultSizeForType(this.drawState.tool).width;
      } else {
        wMm = 40;
      }
    }
    if (hMm < MIN_ELEMENT_SIZE) {
      if (this.drawState?.tool === 'text' || this.drawState?.tool === 'field') {
        hMm = this._defaultSizeForType(this.drawState.tool).height;
      } else {
        hMm = 10;
      }
    }

    this.saveToHistory();
    const el = this._createElementData(this.drawState.tool, xMm, yMm, wMm, hMm);
    if (el.type === 'text' || el.type === 'field') {
      this._autoFitTextElement(el);
    }
    this.elements.push(el);

    this.drawState = null;

    this.renderElements();
    this.selectElement(el.id);
    this._saveLayoutElements();

    // For image, trigger file picker
    if (el.type === 'image' || el.type === 'logo') {
      this._triggerImageUpload(el.id);
    }

    // Switch back to select after placing
      this.setActiveTool('select');
    }

  _startMarqueeSelect(e) {
    const rect = this.pageCanvas.getBoundingClientRect();
    const xPx = e.clientX - rect.left;
    const yPx = e.clientY - rect.top;
    const box = document.createElement('div');
    box.className = 'selection-marquee';
    box.style.left = xPx + 'px';
    box.style.top = yPx + 'px';
    box.style.width = '0px';
    box.style.height = '0px';
    this.pageCanvas.appendChild(box);
    this.marqueeState = {
      startX: xPx,
      startY: yPx,
      currentX: xPx,
      currentY: yPx,
      box,
      moved: false,
      targetElementId: e.target.closest('.canvas-element')?.dataset.id || null,
    };
    e.preventDefault();
    e.stopPropagation();
  }

  _performMarqueeSelect(e) {
    if (!this.marqueeState) return;
    const rect = this.pageCanvas.getBoundingClientRect();
    const xPx = e.clientX - rect.left;
    const yPx = e.clientY - rect.top;
    const ms = this.marqueeState;
    ms.currentX = xPx;
    ms.currentY = yPx;
    ms.moved = ms.moved || Math.abs(xPx - ms.startX) > 3 || Math.abs(yPx - ms.startY) > 3;
    const x = Math.min(xPx, ms.startX);
    const y = Math.min(yPx, ms.startY);
    const w = Math.abs(xPx - ms.startX);
    const h = Math.abs(yPx - ms.startY);
    ms.box.style.left = x + 'px';
    ms.box.style.top = y + 'px';
    ms.box.style.width = w + 'px';
    ms.box.style.height = h + 'px';
  }

  _finishMarqueeSelect(e) {
    const ms = this.marqueeState;
    if (!ms) return;
    const box = ms.box;
    if (box) box.remove();
    this.marqueeState = null;

    if (!ms.moved) {
      this.suppressNextContextMenu = false;
      if (ms.targetElementId) this.selectElement(ms.targetElementId);
      else this.deselectAll();
      return;
    }

    this.suppressNextContextMenu = false;
    const left = this._pxToMm(Math.min(ms.startX, ms.currentX));
    const top = this._pxToMm(Math.min(ms.startY, ms.currentY));
    const right = this._pxToMm(Math.max(ms.startX, ms.currentX));
    const bottom = this._pxToMm(Math.max(ms.startY, ms.currentY));
    const ids = this.elements
      .filter(el => el.x < right && el.x + el.width > left && el.y < bottom && el.y + el.height > top)
      .map(el => el.id);
    this.selectElements(ids);
    e.preventDefault();
    e.stopPropagation();
  }

    // ===== Element Creation =====
  createElement(type, x, y) {
    return this._createElementData(type, x, y);
  }

  _createElementData(type, x, y, width, height) {
    const defaults = this._defaultSizeForType(type);
    width = width || defaults.width;
    height = height || defaults.height;

    const el = {
      id: 'el-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 5),
      type,
      x: Math.max(0, x),
      y: Math.max(0, y),
      width,
      height,
      content: type === 'text' ? 'Text' : '',
      fieldName: '',
      imageData: '',
      lineDirection: 'horizontal',
      datetimeFormat: 'DD/MM/YYYY',
      pagenumFormat: 'Page {n}',
      barcode: { type: 'CODE128', showText: true, fontSize: 10, textColor: '#000000' },
      style: {
        ...this.defaultStyle,
        backgroundColor: 'transparent',
        borderWidth: (type === 'rect' || type === 'table') ? 1 : 0,
        borderColor: '#000000',
        borderStyle: 'solid',
        opacity: 1,
      },
      table: {
        rows: 2,
        cols: 4,
        cells: [],
        theme: 'plain',
        borderMode: 'all',
        colWidths: [1, 1, 1, 1],
        rowHeights: [1, 1],
        detailMode: false,
        colProps: [],
        headerBg: '#ffffff',
        headerColor: '#000000',
        rowBg: '#ffffff',
        altRowBg: '#ffffff',
        footerEnabled: false,
        footerRowHeight: 1,
      },
    };
    if ((type === 'text' || type === 'field') && width === defaults.width && height === defaults.height) {
      this._autoFitTextElement(el);
    }
    return el;
  }

  _defaultSizeForType(type) {
    switch (type) {
      case 'text':     return { width: 10, height: 6 };
      case 'field':    return { width: 10, height: 6 };
      case 'user':     return { width: 45, height: 10 };
      case 'image':
      case 'logo':     return { width: 40, height: 30 };
        case 'rect':     return { width: 60, height: 30 };
      case 'line':     return { width: 80, height: 4 };
      case 'table':    return { width: 140, height: 28 };
      case 'datetime': return { width: 50, height: 8 };
      case 'pagenum':  return { width: 30, height: 8 };
      case 'barcode':  return { width: 45, height: 22 };
      default:         return { width: 60, height: 20 };
    }
  }

  _estimateTextWidthMm(text, fontSizePt) {
    const value = String(text || '');
    const sizePt = parseFloat(fontSizePt) || 12;
    const sizePx = (sizePt * 96) / 72;
    if (!this._measureCanvas) this._measureCanvas = document.createElement('canvas');
    const ctx = this._measureCanvas.getContext('2d');
    if (!ctx) return Math.max(1, value.length) * (sizePt * 0.22);
    ctx.font = `${sizePx}px Arial`;
    const px = Math.max(1, ctx.measureText(value).width);
    return px / MM_TO_PX;
  }

  _autoFitTextElement(el) {
    if (!el || (el.type !== 'text' && el.type !== 'field')) return;
    const style = el.style || {};
    const fontSize = parseFloat(style.fontSize) || 12;
    const displayText = el.type === 'field'
      ? `{${String(el.fieldName || '') || 'Field'}}`
      : String(el.content || 'Text');
    const contentW = this._estimateTextWidthMm(displayText, fontSize);
    const desiredW = Math.max(6, Math.min(120, contentW + 1.6));
    const desiredH = Math.max(4.5, (fontSize * 0.3528) + 1.2);
    el.width = desiredW;
    el.height = desiredH;
  }

  // ===== Selection =====
  _applySelectionToDOM() {
    const selected = new Set(this.selectedIds || []);
    this.pageCanvas.querySelectorAll('.canvas-element').forEach(el => {
      const isSelected = selected.has(el.dataset.id);
      el.classList.toggle('selected', isSelected);
      el.classList.toggle('multi-selected', isSelected && selected.size > 1);
    });
    this._removeResizeHandles();
    if (this.selectedId && selected.size <= 1) {
      const domEl = this.pageCanvas.querySelector(`[data-id="${this.selectedId}"]`);
      if (domEl) this._renderResizeHandles(domEl, this.selectedId);
    }
  }

  selectElement(id) {
    this.selectedId = id;
    this.selectedIds = id ? [id] : [];
    if (!id) {
      this.selectedCell = null;
      this.pageCanvas.querySelectorAll('.el-table-cell').forEach(cell => {
        cell.style.outline = 'none';
      });
    }
    this._applySelectionToDOM();

    if (!id) {
      this._showNoSelectionPanel();
      return;
    }

    this._showElementProperties(id);
  }

  selectElements(ids) {
    const cleanIds = [...new Set((ids || []).filter(id => this._findElement(id)))];
    this.selectedIds = cleanIds;
    this.selectedId = cleanIds[0] || null;
    this.selectedCell = null;
    this.pageCanvas.querySelectorAll('.el-table-cell').forEach(cell => {
      cell.style.outline = 'none';
    });
    this._applySelectionToDOM();
    if (!this.selectedId) {
      this._showNoSelectionPanel();
    } else {
      this._showElementProperties(this.selectedId);
    }
  }

  deselectAll() {
    this.selectElement(null);
  }

  // ===== Move / Drag =====
  _onElementMouseDown(e, elementId) {
    if (e.button !== 0) return;
    if (this.activeTool !== 'select') return;
    e.stopPropagation();

    // Don't start drag if clicking a resize handle, col-resize handle, or inline editor
    if (e.target.classList.contains('resize-handle')) return;
    if (e.target.classList.contains('col-resize-handle')) return;
    if (e.target.classList.contains('row-resize-handle')) return;
    if (e.target.classList.contains('inline-editor')) return;

    if (!this.selectedIds.includes(elementId)) {
      this.selectElement(elementId);
    }
    this.startDrag(e, elementId);
  }

  startDrag(e, elementId) {
    const el = this._findElement(elementId);
    if (!el) return;

    const dragIds = this.selectedIds.includes(elementId) ? this.selectedIds.slice() : [elementId];
    const startPositions = {};
    dragIds.forEach(id => {
      const item = this._findElement(id);
      if (item) startPositions[id] = { x: item.x, y: item.y, width: item.width, height: item.height };
    });

    this.dragState = {
      elementId,
      elementIds: dragIds,
      startPositions,
      startMouseX: e.clientX,
      startMouseY: e.clientY,
      startElX: el.x,
      startElY: el.y,
      moved: false,
    };
    e.preventDefault();
  }

  _showAlignGuides(movingEl) {
    this._hideAlignGuides();
    const SNAP_THRESHOLD = 3; // mm
    const guides = [];
    const mx = movingEl.x, my = movingEl.y, mw = movingEl.width, mh = movingEl.height;
    const mCX = mx + mw/2, mCY = my + mh/2, mR = mx + mw, mB = my + mh;

    this.elements.forEach(el => {
      if (el.id === movingEl.id) return;
      const ex = el.x, ey = el.y, ew = el.width, eh = el.height;
      const eCX = ex + ew/2, eCY = ey + eh/2, eR = ex + ew, eB = ey + eh;

      // Vertical guides (x-axis alignment)
      [[mx, ex], [mx, eCX], [mx, eR], [mCX, ex], [mCX, eCX], [mCX, eR], [mR, ex], [mR, eCX], [mR, eR]].forEach(([a, b]) => {
        if (Math.abs(a - b) < SNAP_THRESHOLD) {
          guides.push({ type: 'v', x: b, y: Math.min(my, ey), h: Math.abs(mB - eB) + Math.max(mh, eh) });
        }
      });
      // Horizontal guides (y-axis alignment)
      [[my, ey], [my, eCY], [my, eB], [mCY, ey], [mCY, eCY], [mCY, eB], [mB, ey], [mB, eCY], [mB, eB]].forEach(([a, b]) => {
        if (Math.abs(a - b) < SNAP_THRESHOLD) {
          guides.push({ type: 'h', y: b, x: Math.min(mx, ex), w: Math.abs(mR - eR) + Math.max(mw, ew) });
        }
      });
    });

    // Remove duplicates
    const seen = new Set();
    guides.forEach(g => {
      const key = `${g.type}${g.type==='v'?g.x:g.y}`;
      if (!seen.has(key)) {
        seen.add(key);
        const line = document.createElement('div');
        line.className = 'align-guide-line';
        const { pageWidthPx, pageHeightPx } = this._getPageDimensions();
        if (g.type === 'v') {
          line.style.cssText = `position:absolute;left:${this._mmToPx(g.x)}px;top:0;width:1px;height:${pageHeightPx}px;background:rgba(255,80,80,0.8);pointer-events:none;z-index:500;`;
        } else {
          line.style.cssText = `position:absolute;top:${this._mmToPx(g.y)}px;left:0;height:1px;width:${pageWidthPx}px;background:rgba(80,80,255,0.8);pointer-events:none;z-index:500;`;
        }
        this.pageCanvas.appendChild(line);
      }
    });
  }

  _hideAlignGuides() {
    this.pageCanvas.querySelectorAll('.align-guide-line').forEach(g => g.remove());
  }

  _performDrag(e) {
    if (!this.dragState) return;
    const ds = this.dragState;
    const dxPx = e.clientX - ds.startMouseX;
    const dyPx = e.clientY - ds.startMouseY;
    if (Math.abs(dxPx) < 1 && Math.abs(dyPx) < 1) return;
    ds.moved = true;

    const dxMm = this._pxToMm(dxPx);
    const dyMm = this._pxToMm(dyPx);

    const { pageWidthMm, pageHeightMm } = this._getPageDimensions();
    const ids = ds.elementIds || [ds.elementId];
    const starts = ds.startPositions || {};
    const bounds = ids.reduce((acc, id) => {
      const start = starts[id];
      if (!start) return acc;
      acc.left = Math.min(acc.left, start.x);
      acc.top = Math.min(acc.top, start.y);
      acc.right = Math.max(acc.right, start.x + start.width);
      acc.bottom = Math.max(acc.bottom, start.y + start.height);
      return acc;
    }, { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity });
    const clampedDxMm = Number.isFinite(bounds.left)
      ? Math.max(-bounds.left, Math.min(pageWidthMm - bounds.right, dxMm))
      : dxMm;
    const clampedDyMm = Number.isFinite(bounds.top)
      ? Math.max(-bounds.top, Math.min(pageHeightMm - bounds.bottom, dyMm))
      : dyMm;

    ids.forEach(id => {
      const el = this._findElement(id);
      const start = starts[id];
      if (!el || !start) return;
      el.x = start.x + clampedDxMm;
      el.y = start.y + clampedDyMm;

      const domEl = this.pageCanvas.querySelector(`[data-id="${id}"]`);
      if (domEl) {
        domEl.style.left = this._mmToPx(el.x) + 'px';
        domEl.style.top = this._mmToPx(el.y) + 'px';
      }
    });

    // Update position fields in properties panel
    const el = this._findElement(ds.elementId);
    const propX = document.getElementById('prop-x');
    const propY = document.getElementById('prop-y');
    if (el && propX && document.getElementById('element-properties')?.classList.contains('hidden') === false) {
      propX.value = el.x.toFixed(1);
      propY.value = el.y.toFixed(1);
    }

    if (el) this._showAlignGuides(el);
  }

  _finishDrag() {
    if (!this.dragState) return;
    if (this.dragState.moved) {
      this.saveToHistory();
      this._saveLayoutElements();
    }
    this.dragState = null;
    this._hideAlignGuides();
  }

  _moveSelectedBy(dxMm, dyMm) {
    const ids = this.selectedIds.length ? this.selectedIds : (this.selectedId ? [this.selectedId] : []);
    if (!ids.length) return;

    const { pageWidthMm, pageHeightMm } = this._getPageDimensions();
    const bounds = ids.reduce((acc, id) => {
      const el = this._findElement(id);
      if (!el) return acc;
      acc.left = Math.min(acc.left, el.x);
      acc.top = Math.min(acc.top, el.y);
      acc.right = Math.max(acc.right, el.x + el.width);
      acc.bottom = Math.max(acc.bottom, el.y + el.height);
      return acc;
    }, { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity });
    if (!Number.isFinite(bounds.left)) return;

    const moveX = Math.max(-bounds.left, Math.min(pageWidthMm - bounds.right, dxMm));
    const moveY = Math.max(-bounds.top, Math.min(pageHeightMm - bounds.bottom, dyMm));
    if (moveX === 0 && moveY === 0) return;

    this.saveToHistory();
    ids.forEach(id => {
      const el = this._findElement(id);
      if (!el) return;
      el.x += moveX;
      el.y += moveY;
    });
    this.renderElements();
    this.selectElements(ids);
    this._saveLayoutElements();
  }

  // ===== Resize =====
  startResize(e, elementId, handle) {
    const el = this._findElement(elementId);
    if (!el) return;
    const rect = this.pageCanvas.getBoundingClientRect();
    this.resizeState = {
      elementId,
      handle,
      startMouseX: e.clientX,
      startMouseY: e.clientY,
      startElX: el.x,
      startElY: el.y,
      startElW: el.width,
      startElH: el.height,
    };
    e.preventDefault();
    e.stopPropagation();
  }

  _performResize(e) {
    if (!this.resizeState) return;
    const rs = this.resizeState;
    const dxPx = e.clientX - rs.startMouseX;
    const dyPx = e.clientY - rs.startMouseY;
    const dxMm = this._pxToMm(dxPx);
    const dyMm = this._pxToMm(dyPx);

    const el = this._findElement(rs.elementId);
    if (!el) return;

    let { startElX: x, startElY: y, startElW: w, startElH: h } = rs;
    const minMm = MIN_ELEMENT_SIZE;

    const h_dir = rs.handle;
    if (h_dir.includes('e')) { w = Math.max(minMm, w + dxMm); }
    if (h_dir.includes('s')) { h = Math.max(minMm, h + dyMm); }
    if (h_dir.includes('w')) {
      const newW = Math.max(minMm, w - dxMm);
      x = x + (w - newW);
      w = newW;
    }
    if (h_dir.includes('n')) {
      const newH = Math.max(minMm, h - dyMm);
      y = y + (h - newH);
      h = newH;
    }

    const { pageWidthMm, pageHeightMm } = this._getPageDimensions();
    el.x = Math.max(0, x);
    el.y = Math.max(0, y);
    el.width = Math.min(w, pageWidthMm - el.x);
    el.height = Math.min(h, pageHeightMm - el.y);

    // If it's a table and width/height changed, scale columns/rows proportionally
    if (el.type === 'table') {
      if (rs.startElW > 0 && el.width !== rs.startElW) {
        const ratio = el.width / rs.startElW;
        el.table.colWidths = this._colWidthsMm(el.table, rs.startElW).map(w => Math.max(2, w * ratio));
      }
      if (rs.startElH > 0 && el.height !== rs.startElH) {
        const ratio = el.height / rs.startElH;
        el.table.rowHeights = this._baseRowHeightsMm(el.table, rs.startElH).map(h => Math.max(2, h * ratio));
        if (el.table.footerEnabled && Number.isFinite(el.table.footerRowHeight)) {
          el.table.footerRowHeight = Math.max(2, el.table.footerRowHeight * ratio);
        }
      }
    }

    // Update DOM
    const domEl = this.pageCanvas.querySelector(`[data-id="${rs.elementId}"]`);
    if (domEl) {
      domEl.style.left = this._mmToPx(el.x) + 'px';
      domEl.style.top = this._mmToPx(el.y) + 'px';
      domEl.style.width = this._mmToPx(el.width) + 'px';
      domEl.style.height = this._mmToPx(el.height) + 'px';
    }

    // Update props
    const pW = document.getElementById('prop-w');
    const pH = document.getElementById('prop-h');
    const pX = document.getElementById('prop-x');
    const pY = document.getElementById('prop-y');
    if (pW) { pW.value = el.width.toFixed(1); pH.value = el.height.toFixed(1); pX.value = el.x.toFixed(1); pY.value = el.y.toFixed(1); }
  }

  _finishResize() {
    if (!this.resizeState) return;
    this.saveToHistory();
    this._saveLayoutElements();
    this.resizeState = null;
  }

  // ===== Double-click inline edit =====
  _onElementDblClick(e, elementId) {
    const el = this._findElement(elementId);
    if (!el) return;
    if (el.type !== 'text' && el.type !== 'field') return;

    const domEl = this.pageCanvas.querySelector(`[data-id="${elementId}"]`);
    if (!domEl) return;

    // Remove existing inline editor
    const existing = domEl.querySelector('.inline-editor');
    if (existing) return;

    const editor = document.createElement('textarea');
    editor.className = 'inline-editor';
    editor.value = el.type === 'text' ? (el.content || '') : (el.fieldName || '');
    editor.style.font = `${el.style?.fontStyle || 'normal'} ${el.style?.fontWeight || 'normal'} ${el.style?.fontSize || 12}pt ${el.style?.fontFamily || 'Arial'}`;
    editor.style.color = el.style?.color || '#000000';
    editor.style.textAlign = el.style?.textAlign || 'left';
    editor.style.padding = '1px';
    editor.style.backgroundColor = 'rgba(255,255,255,0.85)';

    domEl.appendChild(editor);
    editor.focus();
    editor.select();

    const finishEdit = () => {
      const val = editor.value.trim();
      if (el.type === 'text') {
        el.content = val || 'Text';
        this._autoFitTextElement(el);
      } else if (el.type === 'field') {
        el.fieldName = val;
        this._autoFitTextElement(el);
      }
      editor.remove();
      this.saveToHistory();
      const singleDom = this.pageCanvas.querySelector(`[data-id="${elementId}"]`);
      if (singleDom) {
        if (el.type === 'text') singleDom.textContent = el.content;
        else singleDom.textContent = el.fieldName ? `{${el.fieldName}}` : '{Field}';
        singleDom.style.width = this._mmToPx(el.width) + 'px';
        singleDom.style.height = this._mmToPx(el.height) + 'px';
        this._renderResizeHandles(singleDom, elementId);
      }
      this._saveLayoutElements();
    };

    editor.addEventListener('blur', finishEdit);
    editor.addEventListener('keydown', (ev) => {
      if (ev.key === 'Escape') { editor.blur(); }
      if (ev.key === 'Enter' && !ev.shiftKey) { ev.preventDefault(); editor.blur(); }
    });
    e.stopPropagation();
  }

  // ===== Context Menu =====
  _onElementContextMenu(e, elementId) {
    if (this.suppressNextContextMenu) {
      e.preventDefault();
      e.stopPropagation();
      this.suppressNextContextMenu = false;
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    this.selectElement(elementId);
    this._toggleTableColumnContextActions();

    const menu = document.getElementById('context-menu');
    menu.style.left = e.clientX + 'px';
    menu.style.top = e.clientY + 'px';
    menu.classList.remove('hidden');

    // Ensure menu is within viewport
    const mr = menu.getBoundingClientRect();
    if (mr.right > window.innerWidth) menu.style.left = (e.clientX - mr.width) + 'px';
    if (mr.bottom > window.innerHeight) menu.style.top = (e.clientY - mr.height) + 'px';
  }

  _initContextMenu() {
    const menu = document.getElementById('context-menu');
    menu.addEventListener('click', (e) => {
      const item = e.target.closest('[data-action]');
      if (!item) return;
      const action = item.dataset.action;
      this._hideContextMenu();
      if (action === 'duplicate') this.duplicateSelected();
      if (action === 'bring-front') this.bringToFront();
      if (action === 'send-back') this.sendToBack();
      if (action === 'insert-col-before' && this.selectedCell?.isColumn) {
        this.insertTableColumn(this.selectedCell.elementId, this.selectedCell.col, 'before');
      }
      if (action === 'insert-col-after' && this.selectedCell?.isColumn) {
        this.insertTableColumn(this.selectedCell.elementId, this.selectedCell.col, 'after');
      }
      if (action === 'delete' && this.selectedCell?.isColumn) {
        this.deleteTableColumn(this.selectedCell.elementId, this.selectedCell.col);
      } else if (action === 'delete') {
        this.deleteSelected();
      }
    });
  }

  _toggleTableColumnContextActions() {
    const selectedEl = this.selectedId ? this._findElement(this.selectedId) : null;
    const showForColumn = !!(
      this.selectedCell &&
      this.selectedCell.isColumn &&
      selectedEl &&
      selectedEl.type === 'table' &&
      this.selectedCell.elementId === selectedEl.id
    );
    document.querySelectorAll('#context-menu .ctx-table-col').forEach(item => {
      item.classList.toggle('hidden', !showForColumn);
    });
  }

  _hideContextMenu() {
    const menu = document.getElementById('context-menu');
    if (menu) menu.classList.add('hidden');
  }

  insertTableColumn(elementId, colIndex, position = 'after') {
    const el = this._findElement(elementId);
    if (!el || el.type !== 'table') return;

    el.table = el.table || {};
    const currentCols = Math.max(1, el.table.cols || 1);
    const targetCol = Math.max(0, Math.min(currentCols - 1, colIndex || 0));
    const insertAt = position === 'before' ? targetCol : targetCol + 1;

    this.saveToHistory();

    const colWidths = Array.isArray(el.table.colWidths) && el.table.colWidths.length === currentCols
      ? el.table.colWidths.slice()
      : Array(currentCols).fill(1);
    const sourceWidth = colWidths[targetCol] || 1;
    colWidths.splice(insertAt, 0, sourceWidth);
    el.table.colWidths = colWidths;

    const cells = Array.isArray(el.table.cells) ? el.table.cells : [];
    cells.forEach(cell => {
      if (cell.col >= insertAt) cell.col += 1;
    });
    el.table.cells = cells;

    const colProps = Array.isArray(el.table.colProps) ? el.table.colProps.slice() : [];
    while (colProps.length < currentCols) colProps.push({});
    colProps.splice(insertAt, 0, {});
    el.table.colProps = colProps;

    el.table.cols = currentCols + 1;

    this.renderElements();
    this.selectElement(elementId);
    this._selectTableColumn(elementId, insertAt);
    this._saveLayoutElements();
  }

  deleteTableColumn(elementId, colIndex) {
    const el = this._findElement(elementId);
    if (!el || el.type !== 'table') return;

    el.table = el.table || {};
    const currentCols = Math.max(1, el.table.cols || 1);
    if (currentCols <= 1) return;
    const deleteAt = Math.max(0, Math.min(currentCols - 1, colIndex || 0));

    this.saveToHistory();

    const colWidths = Array.isArray(el.table.colWidths) ? el.table.colWidths.slice() : Array(currentCols).fill(1);
    if (colWidths.length < currentCols) {
      while (colWidths.length < currentCols) colWidths.push(1);
    }
    colWidths.splice(deleteAt, 1);
    el.table.colWidths = colWidths;

    const cells = Array.isArray(el.table.cells) ? el.table.cells : [];
    const shiftedCells = [];
    cells.forEach(cell => {
      if (cell.col === deleteAt) return;
      if (cell.col > deleteAt) cell.col -= 1;
      shiftedCells.push(cell);
    });
    el.table.cells = shiftedCells;

    const colProps = Array.isArray(el.table.colProps) ? el.table.colProps.slice() : [];
    while (colProps.length < currentCols) colProps.push({});
    colProps.splice(deleteAt, 1);
    el.table.colProps = colProps;

    el.table.cols = currentCols - 1;

    this.renderElements();
    this.selectElement(elementId);
    this._selectTableColumn(elementId, Math.max(0, deleteAt - 1));
    this._saveLayoutElements();
  }

  // ===== Element operations =====
  deleteSelected() {
    const ids = this.selectedIds.length ? this.selectedIds : (this.selectedId ? [this.selectedId] : []);
    if (!ids.length) return;
    this.saveToHistory();
    this.elements = this.elements.filter(el => !ids.includes(el.id));
    this.selectedId = null;
    this.selectedIds = [];
    this.renderElements();
    this._showNoSelectionPanel();
    this._saveLayoutElements();
  }

  duplicateSelected() {
    const ids = this.selectedIds.length ? this.selectedIds : (this.selectedId ? [this.selectedId] : []);
    if (!ids.length) return;
    this.saveToHistory();
    const copies = ids
      .map(id => this._findElement(id))
      .filter(Boolean)
      .map(original => {
        const copy = JSON.parse(JSON.stringify(original));
        copy.id = 'el-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 5);
        copy.x += 5;
        copy.y += 5;
        return copy;
      });
    this.elements.push(...copies);
    this.renderElements();
    this.selectElements(copies.map(copy => copy.id));
    this._saveLayoutElements();
  }

  bringToFront() {
    if (!this.selectedId) return;
    this.saveToHistory();
    const idx = this.elements.findIndex(el => el.id === this.selectedId);
    if (idx < 0 || idx === this.elements.length - 1) return;
    const [el] = this.elements.splice(idx, 1);
    this.elements.push(el);
    this.renderElements();
    this.selectElement(this.selectedId);
    this._saveLayoutElements();
  }

  sendToBack() {
    if (!this.selectedId) return;
    this.saveToHistory();
    const idx = this.elements.findIndex(el => el.id === this.selectedId);
    if (idx <= 0) return;
    const [el] = this.elements.splice(idx, 1);
    this.elements.unshift(el);
    this.renderElements();
    this.selectElement(this.selectedId);
    this._saveLayoutElements();
  }

  // ===== Properties Panel =====
  _showNoSelectionPanel() {
    const noMsg = document.getElementById('no-selection-msg');
    const elProps = document.getElementById('element-properties');
    if (noMsg) noMsg.classList.remove('hidden');
    if (elProps) elProps.classList.add('hidden');

    // Populate inline page settings
    const p = this.layout?.page;
    if (p) {
      const sz = document.getElementById('inline-page-size');
      const or = document.getElementById('inline-orientation');
      const mt = document.getElementById('inline-margin-top');
      const mr = document.getElementById('inline-margin-right');
      const mb = document.getElementById('inline-margin-bottom');
      const ml = document.getElementById('inline-margin-left');
      const cw = document.getElementById('inline-custom-width');
      const ch = document.getElementById('inline-custom-height');
      if (sz) sz.value = p.size || 'A4';
      if (or) or.value = p.orientation || 'portrait';
      if (mt) mt.value = p.marginTop ?? 15;
      if (mr) mr.value = p.marginRight ?? 15;
      if (mb) mb.value = p.marginBottom ?? 15;
      if (ml) ml.value = p.marginLeft ?? 15;
      if (cw) cw.value = p.customWidthMm ?? p.customWidth ?? 210;
      if (ch) ch.value = p.customHeightMm ?? p.customHeight ?? 297;
      document.getElementById('inline-custom-size-group')?.classList.toggle('hidden', (p.size || 'A4') !== 'custom');

      // Zone settings
      const hEn = document.getElementById('inline-header-enabled');
      const hHt = document.getElementById('inline-header-height');
      const hPg = document.getElementById('inline-header-pages');
      const hOpts = document.getElementById('inline-header-options');
      if (hEn) hEn.checked = !!p.headerEnabled;
      if (hHt) hHt.value = p.headerHeight ?? 20;
      if (hPg) hPg.value = p.headerPages || 'all';
      if (hOpts) hOpts.classList.toggle('hidden', !p.headerEnabled);

      const fEn = document.getElementById('inline-footer-enabled');
      const fHt = document.getElementById('inline-footer-height');
      const fPg = document.getElementById('inline-footer-pages');
      const fOpts = document.getElementById('inline-footer-options');
      if (fEn) fEn.checked = !!p.footerEnabled;
      if (fHt) fHt.value = p.footerHeight ?? 15;
      if (fPg) fPg.value = p.footerPages || 'all';
      if (fOpts) fOpts.classList.toggle('hidden', !p.footerEnabled);
    }

  }

  _showElementProperties(id) {
    const el = this._findElement(id);
    if (!el) { this._showNoSelectionPanel(); return; }

    document.getElementById('no-selection-msg').classList.add('hidden');
    document.getElementById('element-properties').classList.remove('hidden');

    // Position & size
    document.getElementById('prop-x').value = el.x.toFixed(1);
    document.getElementById('prop-y').value = el.y.toFixed(1);
    document.getElementById('prop-w').value = el.width.toFixed(1);
    document.getElementById('prop-h').value = el.height.toFixed(1);

    const style = el.style || {};
    // Typography
      const hasText = ['text','field','user','table','datetime','pagenum'].includes(el.type);
    document.getElementById('prop-group-typography').classList.toggle('hidden', !hasText);
    if (hasText) {
      document.getElementById('prop-font-family').value = style.fontFamily || 'Arial';
      document.getElementById('prop-font-size').value = style.fontSize || 12;
      document.getElementById('prop-color').value = style.color || '#000000';
      this._setStyleBtnActive('prop-bold', style.fontWeight === 'bold');
      this._setStyleBtnActive('prop-italic', style.fontStyle === 'italic');
      this._setStyleBtnActive('prop-underline', style.textDecoration === 'underline');
      document.querySelectorAll('.align-btn[data-align]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.align === (style.textAlign || 'left'));
      });
    }

    // Content
    const hasContent = ['text','field'].includes(el.type);
    document.getElementById('prop-group-content').classList.toggle('hidden', !hasContent);
    document.getElementById('prop-content-text-wrap').classList.toggle('hidden', el.type !== 'text');
    document.getElementById('prop-field-name-wrap').classList.toggle('hidden', el.type !== 'field');

    // DateTime
    document.getElementById('prop-group-datetime').classList.toggle('hidden', el.type !== 'datetime');
    if (el.type === 'datetime') {
      document.getElementById('prop-datetime-format').value = el.datetimeFormat || 'DD/MM/YYYY';
    }

    // Page Number
    document.getElementById('prop-group-pagenum').classList.toggle('hidden', el.type !== 'pagenum');
    if (el.type === 'pagenum') {
      document.getElementById('prop-pagenum-format').value = el.pagenumFormat || 'Page {n}';
    }

    // Barcode
    document.getElementById('prop-group-barcode').classList.toggle('hidden', el.type !== 'barcode');
    if (el.type === 'barcode') {
      const bc = el.barcode || {};
      this._populateBarcodeFieldDropdown(el);
      document.getElementById('prop-barcode-field').value = el.fieldName || '';
      document.getElementById('prop-barcode-static-wrap').classList.toggle('hidden', !!el.fieldName);
      document.getElementById('prop-barcode-value').value = el.content || '';
      document.getElementById('prop-barcode-type').value = bc.type || 'CODE128';
      document.getElementById('prop-barcode-show-text').checked = bc.showText !== false;
      document.getElementById('prop-barcode-font-size').value = bc.fontSize || 10;
      document.getElementById('prop-barcode-text-color').value = bc.textColor || '#000000';
      document.getElementById('prop-barcode-text-opts').classList.toggle('hidden', bc.showText === false);
    }
    if (el.type === 'text') {
      document.getElementById('prop-content').value = el.content || '';
    }
    if (el.type === 'field') {
      this._populateFieldDropdown();
      document.getElementById('prop-field-name').value = el.fieldName || '';
    }

    // Box style (not for line)
    const hasBox = el.type !== 'line';
    document.getElementById('prop-group-box').classList.toggle('hidden', !hasBox);
    if (hasBox) {
      const isBgNone = !style.backgroundColor || style.backgroundColor === 'transparent';
      document.getElementById('prop-bg-none').checked = isBgNone;
      document.getElementById('prop-bg-color').value = isBgNone ? '#ffffff' : style.backgroundColor;
      document.getElementById('prop-bg-color').disabled = isBgNone;
      document.getElementById('prop-opacity').value = style.opacity !== undefined ? style.opacity : 1;
      document.getElementById('prop-opacity-val').textContent = Math.round((style.opacity ?? 1) * 100) + '%';
      document.getElementById('prop-border-width').value = style.borderWidth || 0;
      document.getElementById('prop-border-color').value = style.borderColor || '#000000';
      document.getElementById('prop-border-style').value = style.borderStyle || 'solid';
    }

    // Line
    document.getElementById('prop-group-line').classList.toggle('hidden', el.type !== 'line');
    if (el.type === 'line') {
      document.getElementById('prop-line-direction').value = el.lineDirection || 'horizontal';
      document.getElementById('prop-line-thickness').value = style.borderWidth || 1;
      document.getElementById('prop-line-color').value = style.borderColor || '#000000';
    }

    // Table
    document.getElementById('prop-group-table').classList.toggle('hidden', el.type !== 'table');
    // Hide column props when switching elements (re-shown on header click)
    if (el.type !== 'table') document.getElementById('prop-group-column')?.classList.add('hidden');
    if (el.type === 'table') {
      document.getElementById('prop-table-rows').value = el.table?.rows || 3;
      document.getElementById('prop-table-cols').value = el.table?.cols || 3;
      this._renderTableCellEditor(el);

      // Populate table theme/border controls
      const tbl = el.table || {};
      document.querySelectorAll('.table-theme-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.theme === (tbl.theme || 'plain'));
      });
      const borderModeEl = document.getElementById('prop-table-border-mode');
      if (borderModeEl) borderModeEl.value = tbl.borderMode || 'all';
      const headerBgEl = document.getElementById('prop-table-header-bg');
      if (headerBgEl) headerBgEl.value = tbl.headerBg || '#ffffff';
      const headerColorEl = document.getElementById('prop-table-header-color');
      if (headerColorEl) headerColorEl.value = tbl.headerColor || '#000000';
      const rowBgEl = document.getElementById('prop-table-row-bg');
      if (rowBgEl) rowBgEl.value = tbl.rowBg || '#ffffff';
      const altBgEl = document.getElementById('prop-table-alt-bg');
      if (altBgEl) altBgEl.value = tbl.altRowBg || '#ffffff';
      const detailModeEl = document.getElementById('prop-table-detail-mode');
      if (detailModeEl) detailModeEl.value = tbl.detailMode ? 'detail' : 'static';
      const footerEnabledEl = document.getElementById('prop-table-footer-enabled');
      if (footerEnabledEl) footerEnabledEl.checked = tbl.footerEnabled === true;
    }

    // Image
    document.getElementById('prop-group-image').classList.toggle('hidden', !['image','logo'].includes(el.type));
    if (['image','logo'].includes(el.type) && el.imageData) {
      document.getElementById('prop-image-preview-wrap').classList.remove('hidden');
      document.getElementById('prop-image-preview').src = el.imageData;
    } else {
      document.getElementById('prop-image-preview-wrap').classList.add('hidden');
    }
  }

  _setStyleBtnActive(id, active) {
    const btn = document.getElementById(id);
    if (btn) btn.classList.toggle('active', active);
  }

  _populateBarcodeFieldDropdown(el) {
    const sel = document.getElementById('prop-barcode-field');
    if (!sel) return;
    sel.innerHTML = '<option value="">— Static value —</option>';
    (this.layout.fields || []).forEach(f => {
      const opt = document.createElement('option');
      opt.value = f;
      opt.textContent = f;
      sel.appendChild(opt);
    });
    sel.value = el.fieldName || '';
  }

  _populateFieldDropdown() {
    const sel = document.getElementById('prop-field-name');
    const currentVal = sel.value;
    sel.innerHTML = '<option value="">-- select field --</option>';
    (this.layout.fields || []).forEach(f => {
      const opt = document.createElement('option');
      opt.value = f;
      opt.textContent = f;
      sel.appendChild(opt);
    });
    sel.value = currentVal;
  }

  _renderTableCellEditor(el) {
    const container = document.getElementById('table-cell-visual');
    if (!container) return;
    const rows = el.table?.rows || 3;
    const cols = el.table?.cols || 3;
    const cells = el.table?.cells || [];
    const fields = this.layout.fields || [];

    container.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
    container.style.display = 'grid';
    container.innerHTML = '';

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const cellData = cells.find(cl => cl.row === r && cl.col === c);
        const cell = document.createElement('div');
        cell.className = 'tcv-cell' + (r === 0 ? ' header-cell' : '');
        cell.dataset.row = r;
        cell.dataset.col = c;
        cell.title = `R${r+1}C${c+1} — click to set field`;

        if (cellData?.fieldName) {
          cell.textContent = cellData.fieldName;
          cell.classList.add('has-field');
        } else if (cellData?.content) {
          cell.textContent = cellData.content;
          cell.classList.add('has-text');
        } else {
          cell.textContent = `R${r+1}C${c+1}`;
        }

        // Click to open a mini dropdown
        cell.addEventListener('click', () => {
          // Remove any open dropdowns
          document.querySelectorAll('.tcv-dropdown').forEach(d => d.remove());
          const dd = document.createElement('div');
          dd.className = 'tcv-dropdown';
          dd.style.cssText = `position:absolute;z-index:9999;background:var(--bg-panel);border:1px solid var(--border);border-radius:var(--radius);padding:4px;min-width:120px;box-shadow:var(--shadow-lg);`;
          const input = document.createElement('input');
          input.type = 'text';
          input.placeholder = 'Type text or {FieldName}';
          input.value = cellData?.fieldName ? `{${cellData.fieldName}}` : (cellData?.content || '');
          input.style.cssText = 'width:100%;padding:4px;background:var(--bg-input);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text-primary);font-size:11px;';
          dd.appendChild(input);
          // Field options
          if (fields.length > 0) {
            const opts = document.createElement('div');
            opts.style.cssText = 'margin-top:4px;display:flex;flex-direction:column;gap:2px;max-height:100px;overflow-y:auto;';
            fields.forEach(f => {
              const btn = document.createElement('button');
              btn.textContent = f;
              btn.style.cssText = 'text-align:left;padding:2px 6px;background:none;border:none;color:var(--accent);font-size:11px;cursor:pointer;border-radius:2px;';
              btn.addEventListener('mouseenter', () => btn.style.background = 'var(--bg-hover)');
              btn.addEventListener('mouseleave', () => btn.style.background = 'none');
              btn.addEventListener('click', () => {
                this._updateTableCell(el.id, r, c, f, null);
                dd.remove();
              });
              opts.appendChild(btn);
            });
            dd.appendChild(opts);
          }
          const clearBtn = document.createElement('button');
          clearBtn.textContent = 'Clear';
          clearBtn.style.cssText = 'margin-top:4px;width:100%;padding:2px 6px;background:none;border:1px solid var(--border);color:var(--text-secondary);font-size:11px;cursor:pointer;border-radius:var(--radius-sm);';
          clearBtn.addEventListener('click', () => {
            this._updateTableCell(el.id, r, c, null, '');
            dd.remove();
          });
          dd.appendChild(clearBtn);
          const boldBtn = document.createElement('button');
          boldBtn.textContent = cellData?.style?.fontWeight === 'bold' ? 'Bold: On' : 'Bold: Off';
          boldBtn.style.cssText = 'margin-top:4px;width:100%;padding:2px 6px;background:none;border:1px solid var(--border);color:var(--text-primary);font-size:11px;cursor:pointer;border-radius:var(--radius-sm);font-weight:bold;';
          boldBtn.addEventListener('click', () => {
            this._updateTableCellStyle(el.id, r, c, { fontWeight: cellData?.style?.fontWeight === 'bold' ? 'normal' : 'bold' });
            dd.remove();
          });
          dd.appendChild(boldBtn);
          cell.style.position = 'relative';
          cell.appendChild(dd);
          input.focus();
          input.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
              const val = input.value.trim();
              const fieldRef = val.match(/^\{(.+)\}$/);
              if (fieldRef) this._updateTableCell(el.id, r, c, fieldRef[1], null);
              else this._updateTableCell(el.id, r, c, null, val);
              dd.remove();
            }
            if (ev.key === 'Escape') dd.remove();
          });
          setTimeout(() => {
            document.addEventListener('click', function cleanup(e) {
              if (!dd.contains(e.target) && e.target !== cell) {
                dd.remove();
                document.removeEventListener('click', cleanup);
              }
            });
          }, 50);
        });

        // Drop target
        cell.addEventListener('dragover', (e) => { e.preventDefault(); cell.classList.add('drop-target'); });
        cell.addEventListener('dragleave', () => cell.classList.remove('drop-target'));
        cell.addEventListener('drop', (e) => {
          e.preventDefault();
          cell.classList.remove('drop-target');
          const itemType = e.dataTransfer.getData('application/x-printmore-item-type') || 'field';
          const itemName = e.dataTransfer.getData('text/plain');
          if (itemName) {
            if (itemType === 'text') this._updateTableCell(el.id, r, c, null, itemName);
            else this._updateTableCell(el.id, r, c, itemName, null);
          }
        });

        container.appendChild(cell);
      }
    }
  }

  _updateTableCell(elementId, row, col, fieldName, content) {
    const el = this._findElement(elementId);
    if (!el) return;
    if (!el.table.cells) el.table.cells = [];
    const existing = el.table.cells.find(c => c.row === row && c.col === col);
    if (existing) {
      existing.fieldName = fieldName || '';
      existing.content = content !== null ? content : existing.content;
    } else {
      el.table.cells.push({ row, col, fieldName: fieldName || '', content: content || '' });
    }
    this.saveToHistory();
    this.renderElements();
    this.selectElement(elementId);
    this._saveLayoutElements();
  }

  _updateTableCellStyle(elementId, row, col, stylePatch) {
    const el = this._findElement(elementId);
    if (!el) return;
    if (!el.table.cells) el.table.cells = [];
    let existing = el.table.cells.find(c => c.row === row && c.col === col);
    if (!existing) {
      existing = { row, col, fieldName: '', content: '', style: {} };
      el.table.cells.push(existing);
    }
    existing.style = { ...(existing.style || {}), ...stylePatch };
    this.saveToHistory();
    this.renderElements();
    this.selectElement(elementId);
    this._saveLayoutElements();
  }

  _applySelectedCellStyle(stylePatch) {
    if (!this.selectedCell || this.selectedCell.isColumn) return false;
    this._updateTableCellStyle(this.selectedCell.elementId, this.selectedCell.row, this.selectedCell.col, stylePatch);
    return true;
  }

  // ===== Properties Events =====
  _initPropertiesEvents() {
    const onPosSize = () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el) return;
      el.x = parseFloat(document.getElementById('prop-x').value) || 0;
      el.y = parseFloat(document.getElementById('prop-y').value) || 0;
      el.width = Math.max(1, parseFloat(document.getElementById('prop-w').value) || 10);
      el.height = Math.max(1, parseFloat(document.getElementById('prop-h').value) || 10);
      const domEl = this.pageCanvas.querySelector(`[data-id="${this.selectedId}"]`);
      if (domEl) {
        domEl.style.left = this._mmToPx(el.x) + 'px';
        domEl.style.top = this._mmToPx(el.y) + 'px';
        domEl.style.width = this._mmToPx(el.width) + 'px';
        domEl.style.height = this._mmToPx(el.height) + 'px';
      }
      this._debounceSave();
    };

    ['prop-x','prop-y','prop-w','prop-h'].forEach(id => {
      document.getElementById(id).addEventListener('input', onPosSize);
    });

    // Typography
    const fontChange = () => {
      const ff = document.getElementById('prop-font-family').value;
      const fs = parseFloat(document.getElementById('prop-font-size').value) || 12;
      const fc = document.getElementById('prop-color').value;
      if (this._applySelectedCellStyle({ fontFamily: ff, fontSize: fs, color: fc })) return;
      // Update default style for future elements
      this.defaultStyle.fontFamily = ff;
      this.defaultStyle.fontSize = fs;
      this.defaultStyle.color = fc;
      if (this.selectedId) {
        this.updateElementStyle(this.selectedId, { fontFamily: ff, fontSize: fs, color: fc });
      }
    };
    document.getElementById('prop-font-family').addEventListener('change', fontChange);
    document.getElementById('prop-font-size').addEventListener('input', fontChange);
    document.getElementById('prop-color').addEventListener('input', fontChange);

    document.getElementById('prop-bold').addEventListener('click', () => {
      if (this.selectedCell && !this.selectedCell.isColumn) {
        const el = this._findElement(this.selectedCell.elementId);
        const cell = el?.table?.cells?.find(c => c.row === this.selectedCell.row && c.col === this.selectedCell.col);
        const next = cell?.style?.fontWeight === 'bold' ? 'normal' : 'bold';
        this._applySelectedCellStyle({ fontWeight: next });
        this._setStyleBtnActive('prop-bold', next === 'bold');
        return;
      }
      const newW = (this.defaultStyle.fontWeight === 'bold') ? 'normal' : 'bold';
      this.defaultStyle.fontWeight = newW;
      if (this.selectedId) {
        this.updateElementStyle(this.selectedId, { fontWeight: newW });
      }
      this._setStyleBtnActive('prop-bold', newW === 'bold');
    });
    document.getElementById('prop-italic').addEventListener('click', () => {
      if (this.selectedCell && !this.selectedCell.isColumn) {
        const el = this._findElement(this.selectedCell.elementId);
        const cell = el?.table?.cells?.find(c => c.row === this.selectedCell.row && c.col === this.selectedCell.col);
        const next = cell?.style?.fontStyle === 'italic' ? 'normal' : 'italic';
        this._applySelectedCellStyle({ fontStyle: next });
        this._setStyleBtnActive('prop-italic', next === 'italic');
        return;
      }
      const newS = (this.defaultStyle.fontStyle === 'italic') ? 'normal' : 'italic';
      this.defaultStyle.fontStyle = newS;
      if (this.selectedId) {
        this.updateElementStyle(this.selectedId, { fontStyle: newS });
      }
      this._setStyleBtnActive('prop-italic', newS === 'italic');
    });
    document.getElementById('prop-underline').addEventListener('click', () => {
      if (this.selectedCell && !this.selectedCell.isColumn) {
        const el = this._findElement(this.selectedCell.elementId);
        const cell = el?.table?.cells?.find(c => c.row === this.selectedCell.row && c.col === this.selectedCell.col);
        const next = cell?.style?.textDecoration === 'underline' ? 'none' : 'underline';
        this._applySelectedCellStyle({ textDecoration: next });
        this._setStyleBtnActive('prop-underline', next === 'underline');
        return;
      }
      const newD = (this.defaultStyle.textDecoration === 'underline') ? 'none' : 'underline';
      this.defaultStyle.textDecoration = newD;
      if (this.selectedId) {
        this.updateElementStyle(this.selectedId, { textDecoration: newD });
      }
      this._setStyleBtnActive('prop-underline', newD === 'underline');
    });

    document.querySelectorAll('.align-btn[data-align]').forEach(btn => {
      btn.addEventListener('click', () => {
        const align = btn.dataset.align;
        this.defaultStyle.textAlign = align;
        if (this.selectedId) {
          this.updateElementStyle(this.selectedId, { textAlign: align });
        }
        document.querySelectorAll('.align-btn[data-align]').forEach(b => b.classList.toggle('active', b.dataset.align === align));
      });
    });

    // Heading presets
    const headingPresets = {
      'title': { fontSize: 24, fontWeight: 'bold', fontStyle: 'normal', textDecoration: 'none' },
      'h1':    { fontSize: 18, fontWeight: 'bold', fontStyle: 'normal', textDecoration: 'none' },
      'h2':    { fontSize: 14, fontWeight: 'bold', fontStyle: 'normal', textDecoration: 'none' },
      'h3':    { fontSize: 12, fontWeight: 'bold', fontStyle: 'normal', textDecoration: 'none' },
      'body':  { fontSize: 10, fontWeight: 'normal', fontStyle: 'normal', textDecoration: 'none' },
      'small': { fontSize: 9,  fontWeight: 'normal', fontStyle: 'normal', textDecoration: 'none' },
      'label': { fontSize: 9,  fontWeight: 'bold', fontStyle: 'normal', textDecoration: 'none', color: '#666666' },
    };
    document.querySelectorAll('.preset-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const preset = headingPresets[btn.dataset.preset];
        if (!preset) return;
        // Update default style
        Object.assign(this.defaultStyle, preset);
        // Update UI inputs
        if (preset.fontSize !== undefined) document.getElementById('prop-font-size').value = preset.fontSize;
        if (preset.fontWeight !== undefined) this._setStyleBtnActive('prop-bold', preset.fontWeight === 'bold');
        if (preset.fontStyle !== undefined) this._setStyleBtnActive('prop-italic', preset.fontStyle === 'italic');
        if (preset.textDecoration !== undefined) this._setStyleBtnActive('prop-underline', preset.textDecoration === 'underline');
        if (preset.color !== undefined) document.getElementById('prop-color').value = preset.color;
        // Apply to selected element if any
        if (this.selectedId) {
          this.updateElementStyle(this.selectedId, preset);
        }
      });
    });

    // Box style
    document.getElementById('prop-bg-none').addEventListener('change', () => {
      const isNone = document.getElementById('prop-bg-none').checked;
      document.getElementById('prop-bg-color').disabled = isNone;
      if (!this.selectedId) return;
      this.updateElementStyle(this.selectedId, { backgroundColor: isNone ? 'transparent' : document.getElementById('prop-bg-color').value });
    });
    document.getElementById('prop-bg-color').addEventListener('input', () => {
      if (!this.selectedId) return;
      this.updateElementStyle(this.selectedId, { backgroundColor: document.getElementById('prop-bg-color').value });
    });

    document.getElementById('prop-opacity').addEventListener('input', () => {
      const v = parseFloat(document.getElementById('prop-opacity').value);
      document.getElementById('prop-opacity-val').textContent = Math.round(v * 100) + '%';
      if (!this.selectedId) return;
      this.updateElementStyle(this.selectedId, { opacity: v });
    });

    const borderChange = () => {
      if (!this.selectedId) return;
      this.updateElementStyle(this.selectedId, {
        borderWidth: parseInt(document.getElementById('prop-border-width').value) || 0,
        borderColor: document.getElementById('prop-border-color').value,
        borderStyle: document.getElementById('prop-border-style').value,
      });
    };
    document.getElementById('prop-border-width').addEventListener('input', borderChange);
    document.getElementById('prop-border-color').addEventListener('input', borderChange);
    document.getElementById('prop-border-style').addEventListener('change', borderChange);

    // Content
    document.getElementById('prop-content').addEventListener('input', () => {
      if (!this.selectedId) return;
      this.updateElementContent(this.selectedId, document.getElementById('prop-content').value);
    });

    document.getElementById('prop-field-name').addEventListener('change', () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el) return;
      el.fieldName = document.getElementById('prop-field-name').value;
      this._autoFitTextElement(el);
      const domEl = this.pageCanvas.querySelector(`[data-id="${this.selectedId}"]`);
      if (domEl) {
        domEl.textContent = el.fieldName ? `{${el.fieldName}}` : '{Field}';
        domEl.style.width = this._mmToPx(el.width) + 'px';
        domEl.style.height = this._mmToPx(el.height) + 'px';
      }
      this._renderResizeHandles(domEl, this.selectedId);
      this._debounceSave();
    });

    // Line
    const lineChange = () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el || el.type !== 'line') return;
      el.lineDirection = document.getElementById('prop-line-direction').value;
      this.updateElementStyle(this.selectedId, {
        borderWidth: parseInt(document.getElementById('prop-line-thickness').value) || 1,
        borderColor: document.getElementById('prop-line-color').value,
      });
    };
    document.getElementById('prop-line-direction').addEventListener('change', lineChange);
    document.getElementById('prop-line-thickness').addEventListener('input', lineChange);
    document.getElementById('prop-line-color').addEventListener('input', lineChange);

    // Table
    const tableChange = () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el || el.type !== 'table') return;
      el.table = el.table || {};
      el.table.rows = Math.max(1, parseInt(document.getElementById('prop-table-rows').value) || 2);
      el.table.cols = Math.max(1, parseInt(document.getElementById('prop-table-cols').value) || 4);
      // Reset colWidths / rowHeights when count changes
      const newCols = el.table.cols;
      const newRows = el.table.rows;
      if (!el.table.colWidths || el.table.colWidths.length !== newCols) {
        el.table.colWidths = Array(newCols).fill(1);
      }
      if (!el.table.rowHeights || el.table.rowHeights.length !== newRows) {
        el.table.rowHeights = Array(newRows).fill(1);
      }
      if (!Number.isFinite(el.table.footerRowHeight) || el.table.footerRowHeight <= 0) {
        el.table.footerRowHeight = el.table.rowHeights[Math.max(0, newRows - 1)] || 1;
      }
      this.saveToHistory();
      this.renderElements();
      this.selectElement(this.selectedId);
      this._saveLayoutElements();
    };
    document.getElementById('prop-table-rows').addEventListener('change', tableChange);
    document.getElementById('prop-table-cols').addEventListener('change', tableChange);

    // Table theme buttons
    document.getElementById('table-themes')?.addEventListener('click', (e) => {
      const btn = e.target.closest('.table-theme-btn');
      if (!btn || !this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el || el.type !== 'table') return;
      el.table = el.table || {};
      el.table.theme = btn.dataset.theme;
      // Apply theme defaults
      const themes = {
        'plain':       { headerBg: '#ffffff', headerColor: '#000000', rowBg: '#ffffff', altRowBg: '#ffffff' },
        'dark-header': { headerBg: '#2d2d4e', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#f0f0f8' },
        'blue':        { headerBg: '#2a5298', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#eef3fb' },
        'green':       { headerBg: '#2d6a4f', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#eef7f2' },
        'striped':     { headerBg: '#555555', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#f5f5f5' },
      };
      const t = themes[btn.dataset.theme] || themes['plain'];
      el.table.headerBg = t.headerBg;
      el.table.headerColor = t.headerColor;
      el.table.rowBg = t.rowBg;
      el.table.altRowBg = t.altRowBg;
      document.querySelectorAll('.table-theme-btn').forEach(b => b.classList.toggle('active', b === btn));
      // Update color inputs
      const hBg = document.getElementById('prop-table-header-bg');
      if (hBg) hBg.value = t.headerBg;
      const hClr = document.getElementById('prop-table-header-color');
      if (hClr) hClr.value = t.headerColor;
      const rBg = document.getElementById('prop-table-row-bg');
      if (rBg) rBg.value = t.rowBg;
      const aBg = document.getElementById('prop-table-alt-bg');
      if (aBg) aBg.value = t.altRowBg;
      this.saveToHistory();
      this.renderElements();
      this.selectElement(this.selectedId);
      this._saveLayoutElements();
    });

    const tableStyleChange = () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el || el.type !== 'table') return;
      el.table = el.table || {};
      el.table.borderMode = document.getElementById('prop-table-border-mode')?.value || 'all';
      el.table.headerBg = document.getElementById('prop-table-header-bg')?.value || '#ffffff';
      el.table.headerColor = document.getElementById('prop-table-header-color')?.value || '#000000';
      el.table.rowBg = document.getElementById('prop-table-row-bg')?.value || '#ffffff';
      el.table.altRowBg = document.getElementById('prop-table-alt-bg')?.value || '#ffffff';
      el.table.detailMode = document.getElementById('prop-table-detail-mode')?.value === 'detail';
      el.table.footerEnabled = !!document.getElementById('prop-table-footer-enabled')?.checked;
      if (el.table.footerEnabled && (!Number.isFinite(el.table.footerRowHeight) || el.table.footerRowHeight <= 0)) {
        const base = this._baseRowHeightsMm(el.table, el.height);
        el.table.footerRowHeight = base[Math.max(0, base.length - 1)] || 1;
      }
      this.saveToHistory();
      this.renderElements();
      this.selectElement(this.selectedId);
      this._saveLayoutElements();
    };

    ['prop-table-border-mode','prop-table-header-bg','prop-table-header-color','prop-table-row-bg','prop-table-alt-bg','prop-table-detail-mode','prop-table-footer-enabled'].forEach(id => {
      const elId = document.getElementById(id);
      if (elId) elId.addEventListener('change', tableStyleChange);
      if (elId) elId.addEventListener('input', tableStyleChange);
    });

    // Column properties
    const colPropChange = () => {
      if (!this.selectedCell || !this.selectedCell.isColumn) return;
      const { elementId, col } = this.selectedCell;
      const el = this._findElement(elementId);
      if (!el || el.type !== 'table') return;
      if (!el.table.colProps) el.table.colProps = [];
      while (el.table.colProps.length <= col) el.table.colProps.push({});
      const prev = el.table.colProps[col] || {};

      const activeAlignBtn = document.querySelector('#col-align-btns .align-btn.active[data-col-align]');
      const barcodeEnabled = !!document.getElementById('col-prop-barcode')?.checked;
      el.table.colProps[col] = {
        ...prev,
        textAlign: activeAlignBtn ? activeAlignBtn.dataset.colAlign : 'left',
        paddingLeft: parseInt(document.getElementById('col-prop-pad-left')?.value) || 5,
        paddingRight: parseInt(document.getElementById('col-prop-pad-right')?.value) || 5,
        barcode: barcodeEnabled,
        barcodeType: document.getElementById('col-prop-barcode-type')?.value || 'CODE128',
        barcodeShowText: !!document.getElementById('col-prop-barcode-showtext')?.checked,
        footerType: document.getElementById('col-prop-footer-type')?.value || 'none',
        footerText: document.getElementById('col-prop-footer-text')?.value || '',
        footerMergeNext: !!document.getElementById('col-prop-footer-merge-next')?.checked,
      };

      this.saveToHistory();
      this.renderElements();
      // Re-highlight column after render
      const domEl = this.pageCanvas.querySelector(`[data-id="${elementId}"]`);
      if (domEl) {
        domEl.querySelectorAll(`.el-table-cell[data-col="${col}"]`).forEach(c => {
          c.style.outline = '2px solid var(--accent)';
        });
      }
      this._saveLayoutElements();
    };

    document.querySelectorAll('#col-align-btns .align-btn[data-col-align]').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#col-align-btns .align-btn[data-col-align]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        colPropChange();
      });
    });

    ['col-prop-pad-left', 'col-prop-pad-right'].forEach(id => {
      document.getElementById(id)?.addEventListener('input', colPropChange);
    });

    document.getElementById('col-prop-barcode')?.addEventListener('change', () => {
      const checked = document.getElementById('col-prop-barcode').checked;
      document.getElementById('col-barcode-opts')?.classList.toggle('hidden', !checked);
      colPropChange();
    });
    ['col-prop-barcode-type', 'col-prop-barcode-showtext'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', colPropChange);
    });
    ['col-prop-footer-type', 'col-prop-footer-text', 'col-prop-footer-merge-next'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', colPropChange);
      document.getElementById(id)?.addEventListener('input', colPropChange);
    });

    // Image upload
    document.getElementById('prop-image-upload').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file || !this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        el.imageData = ev.target.result;
        document.getElementById('prop-image-preview-wrap').classList.remove('hidden');
        document.getElementById('prop-image-preview').src = el.imageData;
        this.saveToHistory();
        this.renderElements();
        this.selectElement(this.selectedId);
        this._saveLayoutElements();
      };
      reader.readAsDataURL(file);
    });

    // DateTime format change
    document.getElementById('prop-datetime-format')?.addEventListener('change', () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el || el.type !== 'datetime') return;
      el.datetimeFormat = document.getElementById('prop-datetime-format').value;
      this.renderElements(); this.selectElement(this.selectedId); this._debounceSave();
    });

    // Page number format change
    document.getElementById('prop-pagenum-format')?.addEventListener('change', () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el || el.type !== 'pagenum') return;
      el.pagenumFormat = document.getElementById('prop-pagenum-format').value;
      this.renderElements(); this.selectElement(this.selectedId); this._debounceSave();
    });

    // Barcode changes
    const barcodeChange = () => {
      if (!this.selectedId) return;
      const el = this._findElement(this.selectedId);
      if (!el || el.type !== 'barcode') return;
      const fieldVal = document.getElementById('prop-barcode-field').value;
      el.fieldName = fieldVal;
      el.content = fieldVal ? '' : document.getElementById('prop-barcode-value').value;
      document.getElementById('prop-barcode-static-wrap').classList.toggle('hidden', !!fieldVal);
      el.barcode = {
        type: document.getElementById('prop-barcode-type').value,
        showText: document.getElementById('prop-barcode-show-text').checked,
        fontSize: parseInt(document.getElementById('prop-barcode-font-size').value) || 10,
        textColor: document.getElementById('prop-barcode-text-color').value,
      };
      document.getElementById('prop-barcode-text-opts').classList.toggle('hidden', !el.barcode.showText);
      this.renderElements(); this.selectElement(this.selectedId); this._debounceSave();
    };
    ['prop-barcode-field','prop-barcode-value','prop-barcode-type',
     'prop-barcode-font-size','prop-barcode-text-color'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', barcodeChange);
      document.getElementById(id)?.addEventListener('input', barcodeChange);
    });
    document.getElementById('prop-barcode-show-text')?.addEventListener('change', barcodeChange);

    // Page position buttons
    document.querySelectorAll('.page-pos-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        if (!this.selectedId) return;
        const el = this._findElement(this.selectedId);
        if (!el) return;
        const { pageWidthMm, pageHeightMm } = this._getPageDimensions();
        const p = this.layout.page;
        if (btn.dataset.pos === 'left') {
          el.x = p.marginLeft || 0;
        } else if (btn.dataset.pos === 'center-h') {
          el.x = (pageWidthMm - el.width) / 2;
        } else if (btn.dataset.pos === 'right') {
          el.x = pageWidthMm - el.width - (p.marginRight || 0);
        } else if (btn.dataset.pos === 'center-v') {
          el.y = (pageHeightMm - el.height) / 2;
        }
        this.renderElements();
        this.selectElement(this.selectedId);
        this.saveToHistory();
        this._saveLayoutElements();
      });
    });

    // Action buttons in properties
    document.getElementById('btn-bring-front').addEventListener('click', () => this.bringToFront());
    document.getElementById('btn-send-back').addEventListener('click', () => this.sendToBack());
    document.getElementById('btn-duplicate').addEventListener('click', () => this.duplicateSelected());
    document.getElementById('btn-delete-el').addEventListener('click', () => this.deleteSelected());
  }

  _triggerImageUpload(elementId) {
    const input = document.getElementById('prop-image-upload');
    input.onchange = null;
    // Temporarily reassign
    const handler = (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const el = this._findElement(elementId);
      if (!el) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        el.imageData = ev.target.result;
        this.saveToHistory();
        this.renderElements();
        this.selectElement(elementId);
        this._saveLayoutElements();
      };
      reader.readAsDataURL(file);
      input.removeEventListener('change', handler);
    };
    input.addEventListener('change', handler);
    input.click();
  }

  // ===== Style / Content update =====
  updateElementStyle(id, styleProps) {
    const ids = this.selectedIds.length > 1 && this.selectedIds.includes(id) ? this.selectedIds : [id];
    ids.forEach(itemId => {
      const el = this._findElement(itemId);
      if (!el) return;
      el.style = { ...el.style, ...styleProps };

      const domEl = this.pageCanvas.querySelector(`[data-id="${itemId}"]`);
      if (!domEl) return;

      this._applyStylesToDOM(domEl, el);
    });
    this._applySelectionToDOM();
    this._debounceSave();
  }

  _applyStylesToDOM(domEl, el) {
    const style = el.style || {};
    switch (el.type) {
      case 'text':
        this._applyTextStyle(domEl, style);
        break;
      case 'field':
      case 'user':
        this._applyTextStyle(domEl, style);
        domEl.style.color = style.color || '#3366cc';
        break;
        case 'rect':
          domEl.style.backgroundColor = style.backgroundColor || 'transparent';
        domEl.style.border = style.borderWidth > 0
          ? `${style.borderWidth}px ${style.borderStyle || 'solid'} ${style.borderColor || '#000000'}`
          : 'none';
        break;
      case 'line': {
        domEl.innerHTML = '';
        const inner = document.createElement('div');
        inner.className = 'el-line-inner';
        inner.style.backgroundColor = style.borderColor || '#000000';
        const dir = el.lineDirection || 'horizontal';
        inner.style.width = dir === 'horizontal' ? '100%' : (style.borderWidth || 1) + 'px';
        inner.style.height = dir === 'horizontal' ? (style.borderWidth || 1) + 'px' : '100%';
        domEl.appendChild(inner);
        break;
      }
      case 'table':
        domEl.querySelectorAll('.el-table-cell').forEach(cell => {
          cell.style.fontFamily = style.fontFamily || 'Arial';
          cell.style.fontSize = (style.fontSize || 10) + 'pt';
        });
        break;
      case 'datetime':
        this._applyTextStyle(domEl, style);
        domEl.textContent = this._formatDateTime(el.datetimeFormat || 'DD/MM/YYYY');
        break;
      case 'pagenum':
        this._applyTextStyle(domEl, style);
        { const fmt = el.pagenumFormat || 'Page {n}';
          domEl.textContent = fmt.replace('{n}', '1').replace('{total}', '1'); }
        break;
    }
    domEl.style.opacity = style.opacity !== undefined ? style.opacity : 1;
    if (el.type !== 'line') {
      // Rebuild resize handles
      this._renderResizeHandles(domEl, el.id);
    }
  }

  updateElementContent(id, content) {
    const el = this._findElement(id);
    if (!el) return;
    el.content = content;
    this._autoFitTextElement(el);
    const domEl = this.pageCanvas.querySelector(`[data-id="${id}"]`);
    if (domEl && el.type === 'text') {
      domEl.textContent = content || '';
      domEl.style.width = this._mmToPx(el.width) + 'px';
      domEl.style.height = this._mmToPx(el.height) + 'px';
      this._renderResizeHandles(domEl, id);
    }
    this._debounceSave();
  }

  // ===== Zoom =====
  zoomIn() {
    this.zoom = Math.min(3, this.zoom + 0.1);
    this._applyZoom();
  }

  zoomOut() {
    this.zoom = Math.max(0.3, this.zoom - 0.1);
    this._applyZoom();
  }

  _applyZoom() {
    document.getElementById('zoom-label').textContent = Math.round(this.zoom * 100) + '%';
    this.initCanvas();
    this.renderElements();
  }

  // ===== Grid =====
  toggleGrid() {
    this.showGrid = !this.showGrid;
    this._drawGrid();
    const btn = document.getElementById('btn-grid-toggle');
    if (btn) btn.classList.toggle('active', this.showGrid);
  }

  // ===== Page Settings =====
  openPageSettings() {
    const p = this.layout.page;
    document.getElementById('modal-page-size').value = p.size;
    document.getElementById('modal-orientation').value = p.orientation;
    document.getElementById('modal-custom-width').value = p.customWidthMm ?? p.customWidth ?? 210;
    document.getElementById('modal-custom-height').value = p.customHeightMm ?? p.customHeight ?? 297;
    document.getElementById('modal-custom-size-group')?.classList.toggle('hidden', p.size !== 'custom');
    document.getElementById('modal-margin-top').value = p.marginTop;
    document.getElementById('modal-margin-bottom').value = p.marginBottom;
    document.getElementById('modal-margin-left').value = p.marginLeft;
    document.getElementById('modal-margin-right').value = p.marginRight;
    document.getElementById('modal-page-settings').classList.remove('hidden');
  }

  applyPageSettings() {
    this.layout.page.size = document.getElementById('modal-page-size').value;
    this.layout.page.orientation = document.getElementById('modal-orientation').value;
    if (this.layout.page.size === 'custom') {
      this.layout.page.customWidthMm = Math.max(20, parseFloat(document.getElementById('modal-custom-width').value) || 210);
      this.layout.page.customHeightMm = Math.max(20, parseFloat(document.getElementById('modal-custom-height').value) || 297);
    }
    this.layout.page.marginTop = parseFloat(document.getElementById('modal-margin-top').value) || 0;
    this.layout.page.marginBottom = parseFloat(document.getElementById('modal-margin-bottom').value) || 0;
    this.layout.page.marginLeft = parseFloat(document.getElementById('modal-margin-left').value) || 0;
    this.layout.page.marginRight = parseFloat(document.getElementById('modal-margin-right').value) || 0;
    window.closePageSettingsModal();
    this.saveToHistory();
    this.initCanvas();
    this.renderElements();
    this.saveLayout();
  }

  // ===== History =====
  saveToHistory() {
    // Truncate redo history
    this.history = this.history.slice(0, this.historyIndex + 1);
    this.history.push(JSON.parse(JSON.stringify(this.elements)));
    if (this.history.length > 100) {
      this.history.shift();
      // historyIndex stays at history.length - 1 after shift
    }
    this.historyIndex = this.history.length - 1;
    this.updateUndoRedo();
  }

  undo() {
    if (this.historyIndex <= 0) return;
    this.historyIndex--;
    this.elements = JSON.parse(JSON.stringify(this.history[this.historyIndex]));
    this.selectedId = null;
    this.renderElements();
    this._showNoSelectionPanel();
    this._saveLayoutElements();
    this.updateUndoRedo();
  }

  redo() {
    if (this.historyIndex >= this.history.length - 1) return;
    this.historyIndex++;
    this.elements = JSON.parse(JSON.stringify(this.history[this.historyIndex]));
    this.selectedId = null;
    this.renderElements();
    this._showNoSelectionPanel();
    this._saveLayoutElements();
    this.updateUndoRedo();
  }

  updateUndoRedo() {
    const undoBtn = document.getElementById('btn-undo');
    const redoBtn = document.getElementById('btn-redo');
    if (undoBtn) undoBtn.disabled = this.historyIndex <= 0;
    if (redoBtn) redoBtn.disabled = this.historyIndex >= this.history.length - 1;
  }

  // ===== Serialization =====
  getLayout() {
    return {
      ...this.layout,
      defaultStyle: { ...this.defaultStyle },
      elements: JSON.parse(JSON.stringify(this.elements)),
    };
  }

  loadLayout(layout) {
    this.layout = layout;
    this.elements = JSON.parse(JSON.stringify(layout.elements || []));
    this.initCanvas();
    this.renderElements();
    this.renderFieldsList();
    this.saveToHistory();
  }

  saveLayout() {
    const layout = this.getLayout();
    window.saveLayout(layout);
  }

  _saveLayoutElements() {
    const layout = window.getLayoutById(this.layoutId);
    if (!layout) return;
    layout.elements = JSON.parse(JSON.stringify(this.elements));
    layout.defaultStyle = { ...this.defaultStyle };
    window.saveLayout(layout);
  }

  // ===== Helpers =====
  _findElement(id) {
    return this.elements.find(el => el.id === id) || null;
  }

  _mmToPx(mm) { return mm * MM_TO_PX * this.zoom; }
  _pxToMm(px) { return px / MM_TO_PX / this.zoom; }

  _debounceSave() {
    clearTimeout(this._debounceSaveTimer);
    this._debounceSaveTimer = setTimeout(() => {
      this.saveToHistory();
      this._saveLayoutElements();
    }, 400);
  }

  _onKeyDown(e) {
    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) return;
    const active = document.activeElement;
    if (active && ['INPUT', 'TEXTAREA', 'SELECT'].includes(active.tagName)) return;
    if (this.activeTool !== 'select') return;
    const ids = this.selectedIds.length ? this.selectedIds : (this.selectedId ? [this.selectedId] : []);
    if (!ids.length) return;

    const step = e.shiftKey ? 5 : 1;
    const dx = e.key === 'ArrowLeft' ? -step : (e.key === 'ArrowRight' ? step : 0);
    const dy = e.key === 'ArrowUp' ? -step : (e.key === 'ArrowDown' ? step : 0);
    this._moveSelectedBy(dx, dy);
    e.preventDefault();
  }
}
