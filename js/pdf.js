/**
 * pdf.js — PDF generation for PrintMore
 */

'use strict';

const PDF_MM_TO_PX = 3.7795; // same scale as designer for rendering
const PDF_HD_CANVAS_SCALE = 3;
const PDF_LARGE_JOB_CANVAS_SCALE = 2;

function _sanitizePdfBaseName(name) {
  const cleaned = String(name || 'Layout')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return cleaned || 'Layout';
}

function _pdfTimestampDDMMYYYYHHMMSS() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${pad(now.getDate())}${pad(now.getMonth() + 1)}${now.getFullYear()}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

function _escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Format a date/time string using a format pattern.
 */
function _formatDateTimeStr(format) {
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

/**
 * Replace {FieldName} placeholders in text with actual values.
 */
function applyFieldValues(text, fieldValues) {
  if (!text) return '';
  return text.replace(/\{([^}]+)\}/g, (match, name) => {
    return Object.prototype.hasOwnProperty.call(fieldValues, name) ? (fieldValues[name] || '') : match;
  });
}

/**
 * Determine which zone an element belongs to based on its Y position.
 * Returns 'header', 'footer', or 'body'.
 */
function _getElementZone(el, page, pageHeightMm) {
  const hH = (page.headerEnabled && page.headerHeight > 0) ? (page.headerHeight || 20) : 0;
  const fH = (page.footerEnabled && page.footerHeight > 0) ? (page.footerHeight || 15) : 0;
  // Both zones are measured from physical page top/bottom (same coords as el.y)
  const footerStart = pageHeightMm - fH;
  const elMidY = el.y + el.height / 2;
  if (hH > 0 && elMidY < hH) return 'header';
  if (fH > 0 && elMidY >= footerStart) return 'footer';
  return 'body';
}

function _isHeaderActiveOnPage(page, pageIndex) {
  const pages = page.headerPages || 'all';
  return !!page.headerEnabled && (pages === 'all' || (pages === 'first' && pageIndex === 0));
}

function _isFooterActiveOnPage(page, pageIndex, totalPages) {
  const pages = page.footerPages || 'all';
  return !!page.footerEnabled && (pages === 'all' || (pages === 'last' && pageIndex === totalPages - 1));
}

function _detailStartYForPage(detailEl, page, pageIndex) {
  if (pageIndex === 0) return detailEl.y;
  if (_isHeaderActiveOnPage(page, pageIndex)) return page.headerHeight || 0;
  return page.marginTop ?? 15;
}

/**
 * Render all detail rows in a hidden container, measure each row's actual
 * DOM height, and return page slices with exactly the rows that fit.
 *
 * @returns {Array<Array>} slices — each entry is the detailRows for one page
 */
function _buildPageSlices(detailEl, fieldValues, scale, detailRows, page, hMm) {
  const mb  = page.marginBottom ?? 15;
  const fH = page.footerEnabled ? (page.footerHeight || 0) : 0; // mm
  const footerPages = page.footerPages || 'all';

  // Render ALL rows into a hidden off-screen container to measure heights
  const tmp = document.createElement('div');
  tmp.style.cssText = `position:fixed;left:-9999px;top:0;width:${detailEl.width * scale}px;` +
    `visibility:hidden;pointer-events:none;overflow:visible;`;
  buildTableDOM(tmp, detailEl, fieldValues, scale, detailRows);
  document.body.appendChild(tmp);

  const tblEl = tmp.querySelector('table');
  const trs   = tblEl ? Array.from(tblEl.querySelectorAll('tr')) : [];
  // trs[0] = column-header row, trs[1..n] = data rows
  const rowHeightsPx = trs.map(tr => Math.ceil(tr.getBoundingClientRect().height));
  document.body.removeChild(tmp);

  const headerRowPx = rowHeightsPx[0] || 20; // column-header bar height

  const availablePx = (pageIndex, totalPagesForFooter) => {
    const footerActive = page.footerEnabled && (
      footerPages === 'all' ||
      (footerPages === 'last' && totalPagesForFooter !== null && pageIndex === totalPagesForFooter - 1)
    );
    const startY = _detailStartYForPage(detailEl, page, pageIndex);
    const bottomY = hMm - mb - (footerActive ? fH : 0);
    return Math.max((bottomY - startY) * scale - headerRowPx, 20);
  };

  const slices = [];
  let ri    = 0; // current index into detailRows
  let pageIndex = 0;

  while (ri < detailRows.length) {
    const avail = availablePx(pageIndex, null);
    let usedPx  = 0;
    const batch = [];

    while (ri < detailRows.length) {
      const rh = rowHeightsPx[ri + 1] || headerRowPx; // +1: trs[0] is header
      if (batch.length > 0 && usedPx + rh > avail + 1) break; // +1 rounding buffer
      batch.push(detailRows[ri]);
      usedPx += rh;
      ri++;
    }

    // Guard: always advance at least 1 row to prevent infinite loop
    if (batch.length === 0 && ri < detailRows.length) {
      batch.push(detailRows[ri++]);
    }

    slices.push(batch);
    pageIndex++;
  }

  if (page.footerEnabled && footerPages === 'last' && slices.length > 0) {
    let lastIndex = slices.length - 1;
    let lastRows = slices[lastIndex];
    const lastStartIndex = slices.slice(0, lastIndex).reduce((sum, pageRows) => sum + pageRows.length, 0);
    let lastUsed = lastRows.reduce((sum, _row, rowIndex) => {
      const originalIndex = lastStartIndex + rowIndex;
      return sum + (rowHeightsPx[originalIndex + 1] || headerRowPx);
    }, 0);
    const lastAvail = availablePx(lastIndex, slices.length);

    if (lastUsed > lastAvail && lastRows.length > 1) {
      const moved = [];
      if (lastRows.length > 1) {
        const row = lastRows.pop();
        moved.unshift(row);
      }
      if (moved.length) slices.push(moved);
    }
  }

  return slices.length ? slices : [[]];
}

/**
 * Build a single page's DOM container with specified elements.
 * detailOverride: { el, rows } — if set, renders detail table with these rows (overrides default).
 */
function _buildPageDOM(page, wMm, hMm, elements, fieldValues, scale, detailOverride, pageNum, totalPages) {
  const wPx = wMm * scale;
  const hPx = hMm * scale;

  const container = document.createElement('div');
  container.style.cssText = `position:relative;width:${wPx}px;height:${hPx}px;background:#ffffff;overflow:hidden;font-family:Arial,sans-serif;box-sizing:border-box;`;

  elements.forEach(el => {
    const isDetailEl = detailOverride && el === detailOverride.el;
    const wrapper = document.createElement('div');
    // el.x / el.y are in physical-page mm from the page top-left (same origin as the designer canvas).
    // Do NOT add margins here — the margins are already baked into the stored coordinates.
    const xPx = el.x * scale;
    const yPx = el.y * scale;
    const wElPx = el.width * scale;
    const hElPx = el.height * scale;
    wrapper.style.cssText = `position:absolute;left:${xPx}px;top:${yPx}px;width:${wElPx}px;height:${hElPx}px;box-sizing:border-box;overflow:hidden;opacity:${el.style?.opacity !== undefined ? el.style.opacity : 1};`;

    if (el.type === 'table' && isDetailEl) {
      wrapper.style.height = 'auto';
      wrapper.style.overflow = 'visible';
      buildTableDOM(wrapper, el, fieldValues, scale, detailOverride.rows, detailOverride.allRows || detailOverride.rows);
    } else {
      switch (el.type) {
        case 'text':     buildTextDOM(wrapper, el, fieldValues); break;
        case 'field':    buildFieldDOM(wrapper, el, fieldValues); break;
        case 'user':     buildUserDOM(wrapper, el); break;
        case 'image':
        case 'logo':     buildImageDOM(wrapper, el); break;
        case 'rect':     buildRectDOM(wrapper, el); break;
        case 'line':     buildLineDOM(wrapper, el, scale); break;
        case 'table':    buildTableDOM(wrapper, el, fieldValues, scale, null); break;
        case 'datetime': buildDateTimeDOM(wrapper, el); break;
        case 'pagenum':  buildPageNumDOM(wrapper, el, pageNum || 1, totalPages || 1); break;
        case 'barcode':  buildBarcodeDOM(wrapper, el, fieldValues); break;
      }
    }
    container.appendChild(wrapper);
  });

  return container;
}

/**
 * Build a single element's DOM node for PDF rendering.
 */
function buildElementDOM(el, fieldValues, scale, detailRows) {
  const wrapper = document.createElement('div');
  wrapper.style.cssText = `
    position: absolute;
    left: ${el.x * scale}px;
    top: ${el.y * scale}px;
    width: ${el.width * scale}px;
    height: ${el.height * scale}px;
    box-sizing: border-box;
    overflow: hidden;
    opacity: ${el.style?.opacity !== undefined ? el.style.opacity : 1};
  `;

  switch (el.type) {
    case 'text':
      buildTextDOM(wrapper, el, fieldValues);
      break;
      case 'field':
        buildFieldDOM(wrapper, el, fieldValues);
        break;
      case 'user':
        buildUserDOM(wrapper, el);
        break;
    case 'image':
    case 'logo':
      buildImageDOM(wrapper, el);
      break;
    case 'rect':
      buildRectDOM(wrapper, el);
      break;
    case 'line':
      buildLineDOM(wrapper, el, scale);
      break;
    case 'table':
      buildTableDOM(wrapper, el, fieldValues, scale, detailRows);
      break;
    case 'datetime':
      buildDateTimeDOM(wrapper, el);
      break;
    case 'pagenum':
      buildPageNumDOM(wrapper, el, el._pageNum || 1, el._totalPages || 1);
      break;
    case 'barcode':
      buildBarcodeDOM(wrapper, el, fieldValues);
      break;
    default:
      return null;
  }

  return wrapper;
}

function applyTextStyles(el, domEl, style) {
  domEl.style.fontFamily = style.fontFamily || 'Arial';
  domEl.style.fontSize = (style.fontSize || 12) + 'pt';
  domEl.style.fontWeight = style.fontWeight || 'normal';
  domEl.style.fontStyle = style.fontStyle || 'normal';
  domEl.style.textDecoration = style.textDecoration || 'none';
  domEl.style.color = style.color || '#000000';
  domEl.style.textAlign = style.textAlign || 'left';
  domEl.style.lineHeight = '1.3';
  domEl.style.wordBreak = 'break-word';

  if (style.backgroundColor && style.backgroundColor !== 'transparent') {
    domEl.style.backgroundColor = style.backgroundColor;
  }
  if (style.borderWidth > 0) {
    domEl.style.border = `${style.borderWidth}px ${style.borderStyle || 'solid'} ${style.borderColor || '#000000'}`;
  }
  domEl.style.padding = '1px';
  domEl.style.width = '100%';
  domEl.style.height = '100%';
  domEl.style.display = 'flex';
  domEl.style.alignItems = 'flex-start';
  domEl.style.boxSizing = 'border-box';
}

function buildTextDOM(wrapper, el, fieldValues) {
  const style = el.style || {};
  const inner = document.createElement('div');
  applyTextStyles(el, inner, style);
  // Apply field values even in static text (in case user put {FieldName} in text)
  inner.textContent = applyFieldValues(el.content || '', fieldValues);
  wrapper.appendChild(inner);
}

function buildFieldDOM(wrapper, el, fieldValues) {
  const style = el.style || {};
  const inner = document.createElement('div');
  applyTextStyles(el, inner, style);
  const val = el.fieldName ? (fieldValues[el.fieldName] !== undefined ? fieldValues[el.fieldName] : '') : '';
  inner.textContent = val;
  wrapper.appendChild(inner);
}

function buildUserDOM(wrapper, el) {
  const inner = document.createElement('div');
  inner.style.cssText = 'width:100%;height:100%;display:flex;align-items:center;';
  const align = el.style?.textAlign || 'left';
  inner.style.justifyContent = align === 'center' ? 'center' : (align === 'right' ? 'flex-end' : 'flex-start');
  const user = window.AuthStore?.currentUser?.();
  inner.textContent = user?.username || '';
  wrapper.appendChild(inner);
}

function buildImageDOM(wrapper, el) {
  if (el.imageData) {
    const img = document.createElement('img');
    img.src = el.imageData;
    img.style.cssText = `width:100%;height:100%;object-fit:contain;display:block;`;
    wrapper.appendChild(img);
  } else {
    // Placeholder — grey box
    wrapper.style.backgroundColor = '#f0f0f0';
    wrapper.style.border = '1px dashed #cccccc';
  }
}

function buildRectDOM(wrapper, el) {
  const style = el.style || {};
  if (style.backgroundColor && style.backgroundColor !== 'transparent') {
    wrapper.style.backgroundColor = style.backgroundColor;
  }
  if (style.borderWidth > 0) {
    wrapper.style.border = `${style.borderWidth}px ${style.borderStyle || 'solid'} ${style.borderColor || '#000000'}`;
  }
}

function buildLineDOM(wrapper, el, scale) {
  const style = el.style || {};
  const color = style.borderColor || '#000000';
  const thickness = (style.borderWidth || 1);
  const direction = el.lineDirection || 'horizontal';

  wrapper.style.display = 'flex';
  wrapper.style.alignItems = 'center';
  wrapper.style.justifyContent = 'center';

  const line = document.createElement('div');
  if (direction === 'horizontal') {
    line.style.cssText = `width:100%;height:${thickness}px;background:${color};`;
  } else {
    line.style.cssText = `width:${thickness}px;height:100%;background:${color};`;
  }
  wrapper.appendChild(line);
}

function buildDateTimeDOM(wrapper, el) {
  const style = el.style || {};
  const inner = document.createElement('div');
  applyTextStyles(el, inner, style);
  inner.textContent = _formatDateTimeStr(el.datetimeFormat || 'DD/MM/YYYY');
  wrapper.appendChild(inner);
}

function buildPageNumDOM(wrapper, el, pageNum, totalPages) {
  const style = el.style || {};
  const inner = document.createElement('div');
  applyTextStyles(el, inner, style);
  const fmt = el.pagenumFormat || 'Page {n}';
  inner.textContent = fmt.replace('{n}', pageNum).replace('{total}', totalPages);
  wrapper.appendChild(inner);
}

function buildBarcodeDOM(wrapper, el, fieldValues) {
  const bc = el.barcode || {};
  const rawValue = el.fieldName
    ? (fieldValues[el.fieldName] !== undefined ? String(fieldValues[el.fieldName]) : '')
    : (el.content || '');
  const value = rawValue || '0000000000';

  // Use individual property assignments — never reassign cssText (would reset position/size)
  wrapper.style.background = '#fff';
  wrapper.style.overflow = 'hidden';

  // Render barcode onto a canvas sized to exactly the wrapper dimensions (in px).
  // Avoid display:flex and object-fit:contain — html2canvas does not support either
  // and will misplace or distort the barcode in the PDF output.
  const wPx = el.width  * PDF_MM_TO_PX;
  const hPx = el.height * PDF_MM_TO_PX;
  const HIRES = 4;
  const fontSize   = (bc.fontSize || 10) * HIRES;
  const textAllow  = bc.showText !== false ? fontSize * 1.4 + HIRES * 2 : 0;
  const barHeight  = Math.max(10, Math.round(hPx * HIRES - textAllow - HIRES * 4));

  const canvas = document.createElement('canvas');
  try {
    JsBarcode(canvas, value, {
      format:       bc.type || 'CODE128',
      displayValue: bc.showText !== false,
      fontSize,
      lineColor:    bc.textColor || '#000000',
      height:       barHeight,
      width:        HIRES * 2,
      margin:       HIRES * 2,
      textMargin:   HIRES * 2,
      background:   '#ffffff',
    });
    // Resize canvas to exact wrapper px dimensions so img fills at 1:1 with no scaling needed
    const scaled = document.createElement('canvas');
    scaled.width  = Math.round(wPx);
    scaled.height = Math.round(hPx);
    const ctx = scaled.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, scaled.width, scaled.height);
    // Draw the barcode centred in the scaled canvas
    const srcW = canvas.width, srcH = canvas.height;
    const dstW = scaled.width, dstH = scaled.height;
    const scaleF = Math.min(dstW / srcW, dstH / srcH);
    const drawW  = srcW * scaleF;
    const drawH  = srcH * scaleF;
    const drawX  = (dstW - drawW) / 2;
    const drawY  = (dstH - drawH) / 2;
    ctx.drawImage(canvas, drawX, drawY, drawW, drawH);
    const img = document.createElement('img');
    img.src = scaled.toDataURL('image/png');
    // Exact pixel size — no CSS scaling, no object-fit — html2canvas renders it perfectly
    img.style.cssText = `position:absolute;left:0;top:0;width:${dstW}px;height:${dstH}px;display:block;`;
    wrapper.appendChild(img);
  } catch (e) {
    wrapper.style.border = '1px dashed #ccc';
    const lbl = document.createElement('div');
    lbl.textContent = value;
    lbl.style.cssText = 'padding:4px;font-size:10px;color:#666;';
    wrapper.appendChild(lbl);
  }
}

function buildTableDOM(wrapper, el, fieldValues, scale, detailRows, allDetailRows) {
  const tbl = el.table || { rows: 2, cols: 4, cells: [], theme: 'plain', borderMode: 'all' };
  const isDetail = tbl.detailMode === true;
  const footerEnabled = tbl.footerEnabled === true;
  const cols = tbl.cols || 3;
  const cells = tbl.cells || [];
  const colWidths = tbl.colWidths || null;
  const style = el.style || {};
  const aggregateRows = (allDetailRows && allDetailRows.length) ? allDetailRows : detailRows;

  // Detail mode: row 0 = header, rows 1..N = one per detailRows entry
  const dataRows = (detailRows && detailRows.length > 0) ? detailRows : null;
  const baseRows = isDetail && dataRows
    ? dataRows.length + 1   // header + data rows
    : (tbl.rows || 3);
  const rows = baseRows + (footerEnabled ? 1 : 0);

  const themes = {
    'plain':       { headerBg: '#ffffff', headerColor: '#000000', rowBg: '#ffffff', altRowBg: '#ffffff' },
    'dark-header': { headerBg: '#2d2d4e', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#f0f0f8' },
    'blue':        { headerBg: '#2a5298', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#eef3fb' },
    'green':       { headerBg: '#2d6a4f', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#eef7f2' },
    'striped':     { headerBg: '#555555', headerColor: '#ffffff', rowBg: '#ffffff', altRowBg: '#f5f5f5' },
  };
  const t = themes[tbl.theme || 'plain'] || themes['plain'];
  const headerBg = tbl.headerBg || t.headerBg;
  const headerColor = tbl.headerColor || t.headerColor;
  const rowBg = tbl.rowBg || t.rowBg;
  const altRowBg = tbl.altRowBg || t.altRowBg;
  const borderMode = tbl.borderMode || 'all';
  const bc = style.borderColor || '#cccccc';
  const bw = Math.max(1, style.borderWidth !== undefined ? style.borderWidth : 1);
  const bs = style.borderStyle || 'solid';
  const colProps = tbl.colProps || [];
  const toNumber = (value) => {
    if (value === null || value === undefined) return null;
    const normalized = String(value).replace(/,/g, '').trim();
    if (!normalized) return null;
    const num = Number(normalized);
    return Number.isFinite(num) ? num : null;
  };
  const resolveDetailFieldForCol = (colIdx) => {
    const rowOne = cells.find(cl => cl.row === 1 && cl.col === colIdx);
    if (rowOne?.fieldName) return rowOne.fieldName;
    const anyMapped = cells.find(cl => cl.col === colIdx && cl.row > 0 && cl.fieldName);
    return anyMapped?.fieldName || '';
  };
  const computeFooterValue = (colIdx, cp) => {
    const footerType = cp.footerType || 'none';
    if (footerType === 'text') return cp.footerText || '';
    if (footerType === 'sum') {
      const fieldName = resolveDetailFieldForCol(colIdx);
      if (!fieldName || !Array.isArray(aggregateRows) || !aggregateRows.length) return '';
      const total = aggregateRows.reduce((sum, row) => {
        const n = toNumber(row?.[fieldName]);
        return n === null ? sum : sum + n;
      }, 0);
      if (!Number.isFinite(total)) return '';
      return Number.isInteger(total) ? String(total) : total.toFixed(2);
    }
    return '';
  };

  const tableEl = document.createElement('table');
  tableEl.style.cssText = `width:100%;border-collapse:collapse;table-layout:fixed;font-family:${style.fontFamily || 'Arial'};font-size:${style.fontSize || 10}pt;color:${style.color || '#000000'};`;

  // colgroup for column widths
  if (colWidths && colWidths.length === cols) {
    const cg = document.createElement('colgroup');
    const totalW = colWidths.reduce((s, w) => s + w, 0);
    colWidths.forEach(w => {
      const col = document.createElement('col');
      col.style.width = (w / totalW * 100).toFixed(2) + '%';
      cg.appendChild(col);
    });
    tableEl.appendChild(cg);
  }

  // Compute per-row heights from designer's rowHeights array.
  // rowHeights may be proportional (old, sum < 20) or absolute mm (new).
  const storedRH = tbl.rowHeights && tbl.rowHeights.length >= 2
    ? tbl.rowHeights
    : Array(Math.max(tbl.rows || 3, 2)).fill(1);
  const rhSum = storedRH.reduce((s, h) => s + h, 0);
  // Convert to absolute mm if proportional
  const rowHeightsBaseMm = rhSum < 20
    ? storedRH.map(h => (h / rhSum) * el.height)
    : storedRH.slice();
  const footerRowMm = (Number.isFinite(tbl.footerRowHeight) && tbl.footerRowHeight > 0)
    ? tbl.footerRowHeight
    : (rowHeightsBaseMm[Math.max(1, rowHeightsBaseMm.length) - 1] || rowHeightsBaseMm[0] || 5);
  const rowHeightsMm = footerEnabled ? rowHeightsBaseMm.concat([footerRowMm]) : rowHeightsBaseMm;
  const headerRowPx = rowHeightsMm[0] * scale;
  const dataRowPx   = rowHeightsMm[1] * scale;

  for (let r = 0; r < rows; r++) {
    const tr = document.createElement('tr');
    const isHeaderRow = r === 0;
    const isFooterRow = footerEnabled && r === rows - 1;

    // Apply designer row height so PDF matches the designed layout
    const rowHpx = rowHeightsMm[Math.min(r, rowHeightsMm.length - 1)] * scale;
    tr.style.height = rowHpx + 'px';

    let skipCols = 0;
    for (let c = 0; c < cols; c++) {
      if (skipCols > 0) {
        skipCols--;
        continue;
      }
      const td = document.createElement(isHeaderRow ? 'th' : 'td');
      const cp = colProps[c] || {};
      const mergeNext = isFooterRow && cp.footerMergeNext === true && c < cols - 1;
      if (mergeNext) {
        td.colSpan = 2;
        skipCols = 1;
      }
      const colAlign = cp.textAlign || style.textAlign || 'left';
      const padLeft = cp.paddingLeft !== undefined ? cp.paddingLeft + 'px' : '5px';
      const padRight = cp.paddingRight !== undefined ? cp.paddingRight + 'px' : '5px';
      td.style.paddingTop = '1px';
      td.style.paddingBottom = '1px';
      td.style.paddingLeft = padLeft;
      td.style.paddingRight = padRight;
      td.style.overflow = 'hidden';
      td.style.verticalAlign = 'middle';
      td.style.textAlign = colAlign;

      if (isHeaderRow) {
        td.style.backgroundColor = headerBg;
        td.style.color = headerColor;
        td.style.fontWeight = 'bold';
      } else if (isFooterRow) {
        td.style.backgroundColor = rowBg;
        td.style.color = style.color || '#000000';
        td.style.fontWeight = 'bold';
      } else {
        td.style.backgroundColor = ((r % 2) === 0) ? altRowBg : rowBg;
        td.style.color = style.color || '#000000';
      }

      if (borderMode === 'all') {
        td.style.border = `${bw}px ${bs} ${bc}`;
      } else if (borderMode === 'outer') {
        td.style.border = 'none';
        if (r === 0) td.style.borderTop = `${bw}px ${bs} ${bc}`;
        if (r === rows - 1) td.style.borderBottom = `${bw}px ${bs} ${bc}`;
        if (c === 0) td.style.borderLeft = `${bw}px ${bs} ${bc}`;
        if (c === cols - 1 || (mergeNext && c + 1 === cols - 1)) td.style.borderRight = `${bw}px ${bs} ${bc}`;
      } else if (borderMode === 'header-outer') {
        td.style.border = 'none';
        if (r === 0) td.style.borderBottom = `2px ${bs} ${bc}`;
        if (isFooterRow) td.style.borderTop = `2px ${bs} ${bc}`;
      } else {
        td.style.border = 'none';
      }

      // Find cell definition from designer
      const cellDef = cells.find(cl => cl.row === r && cl.col === c);
      const repeatingCellDef = isDetail && !isHeaderRow && !isFooterRow
        ? (cells.find(cl => cl.row === 1 && cl.col === c) || cellDef)
        : cellDef;
      const effectiveStyle = repeatingCellDef?.style || {};
      if (effectiveStyle.fontFamily) td.style.fontFamily = effectiveStyle.fontFamily;
      if (effectiveStyle.fontSize) td.style.fontSize = effectiveStyle.fontSize + 'pt';
      if (effectiveStyle.fontWeight) td.style.fontWeight = effectiveStyle.fontWeight;
      if (effectiveStyle.fontStyle) td.style.fontStyle = effectiveStyle.fontStyle;
      if (effectiveStyle.textDecoration) td.style.textDecoration = effectiveStyle.textDecoration;
      if (effectiveStyle.color) td.style.color = effectiveStyle.color;

      if (isHeaderRow) {
        // Header row: show cell content or field name as label
        if (cellDef?.content) td.textContent = cellDef.content;
        else if (cellDef?.fieldName) td.textContent = cellDef.fieldName;
        else td.textContent = `Col ${c + 1}`;
      } else if (isFooterRow) {
        td.textContent = computeFooterValue(c, cp);
      } else if (cp.barcode) {
        // Barcode column — resolve the cell value then render as barcode
        let cellValue = '';
        if (isDetail && dataRows) {
          const dataCellDef = cells.find(cl => cl.row === 1 && cl.col === c) || cells.find(cl => cl.row === r && cl.col === c);
          if (dataCellDef?.fieldName && dataRows[r - 1]) cellValue = String(dataRows[r - 1][dataCellDef.fieldName] || '');
          else if (dataCellDef?.content) cellValue = applyFieldValues(dataCellDef.content, fieldValues);
        } else {
          if (cellDef?.fieldName) cellValue = String(fieldValues[cellDef.fieldName] || '');
          else if (cellDef?.content) cellValue = applyFieldValues(cellDef.content, fieldValues);
        }
        if (cellValue && typeof JsBarcode !== 'undefined') {
          td.style.padding = '2px';
          td.style.textAlign = 'center';
          const HIRES = 3;
          const fontSize = 8 * HIRES;
          const textAllow = cp.barcodeShowText !== false ? fontSize * 1.4 + HIRES * 2 : 0;
          const barHeight = Math.max(8, Math.round(rowHpx * HIRES - textAllow - HIRES * 4));
          const canvas = document.createElement('canvas');
          try {
            JsBarcode(canvas, cellValue, {
              format: cp.barcodeType || 'CODE128',
              displayValue: cp.barcodeShowText !== false,
              fontSize, height: barHeight, width: HIRES, margin: HIRES,
              textMargin: HIRES, background: '#ffffff',
              lineColor: '#000000',
            });
            const img = document.createElement('img');
            img.src = canvas.toDataURL('image/png');
            img.style.cssText = 'max-width:100%;height:100%;object-fit:contain;display:block;margin:auto;';
            td.appendChild(img);
          } catch(e) {
            td.textContent = cellValue;
          }
        } else {
          td.textContent = cellValue;
        }
      } else if (isDetail && dataRows) {
        // Detail row: find which field this column maps to (from row=1 cells in designer)
        const dataCellDef = cells.find(cl => cl.row === 1 && cl.col === c) || cells.find(cl => cl.row === r && cl.col === c);
        if (dataCellDef?.fieldName && dataRows[r - 1]) {
          td.textContent = dataRows[r - 1][dataCellDef.fieldName] || '';
        } else if (dataCellDef?.content) {
          td.textContent = applyFieldValues(dataCellDef.content, fieldValues);
        }
      } else {
        // Static mode
        if (cellDef?.fieldName) {
          td.textContent = fieldValues[cellDef.fieldName] !== undefined ? fieldValues[cellDef.fieldName] : '';
        } else if (cellDef?.content) {
          td.textContent = applyFieldValues(cellDef.content, fieldValues);
        }
      }

      tr.appendChild(td);
    }
    tableEl.appendChild(tr);
  }

  // If detail mode, allow overflow so all rows are visible
  if (isDetail) {
    wrapper.style.height = 'auto';
    wrapper.style.overflow = 'visible';
  }

  wrapper.appendChild(tableEl);
}

/**
 * Main PDF generation function — supports multi-page with header/footer zones.
 * @param {Object} layout
 * @param {Object} fieldValues
 * @param {Array}  [detailRows]
 */
async function generatePDF(layout, fieldValues, detailRows) {
  const page = layout.page;
  const sizes = window.PAGE_SIZES;
  let wMm;
  let hMm;
  if (page.size === 'custom') {
    wMm = Math.max(20, parseFloat(page.customWidthMm ?? page.customWidth) || 210);
    hMm = Math.max(20, parseFloat(page.customHeightMm ?? page.customHeight) || 297);
  } else {
    ({ width: wMm, height: hMm } = sizes[page.size] || sizes['A4']);
  }
  if (page.orientation === 'landscape') [wMm, hMm] = [hMm, wMm];
  const scale = PDF_MM_TO_PX;

  const allElements = layout.elements || [];
  const mt = page.marginTop ?? 15;
  const mb = page.marginBottom ?? 15;
  const contentH = hMm - mt - mb; // usable height in mm (inside margins)

  const hEnabled = !!page.headerEnabled;
  const fEnabled = !!page.footerEnabled;
  const hH = hEnabled ? (page.headerHeight || 20) : 0; // mm
  const fH = fEnabled ? (page.footerHeight || 15) : 0; // mm
  const hPages = page.headerPages || 'all';
  const fPages = page.footerPages || 'all';

  // Classify elements by zone
  const headerEls = allElements.filter(el => _getElementZone(el, page, hMm) === 'header');
  const footerEls = allElements.filter(el => _getElementZone(el, page, hMm) === 'footer');
  const bodyEls   = allElements.filter(el => _getElementZone(el, page, hMm) === 'body');

  // Find detail table (first detail-mode table in body)
  const detailEl = bodyEls.find(el => el.type === 'table' && el.table?.detailMode === true);
  const staticBodyEls = bodyEls.filter(el => el !== detailEl);

  const hasData = detailEl && detailRows && detailRows.length > 0;

  // ── Pagination ──────────────────────────────────────────────────────────
  // pageSlices[i] = array of detailRows for page i, or null if no pagination
  let pageSlices;

  if (hasData) {
    pageSlices = _buildPageSlices(detailEl, fieldValues, scale, detailRows, page, hMm);
  } else {
    pageSlices = [null]; // single page, no detail pagination
  }

  const totalPages = pageSlices.length;
  const detailCount = detailRows?.length || 0;
  const canvasScale = (totalPages > 12 || detailCount > 250)
    ? PDF_LARGE_JOB_CANVAS_SCALE
    : PDF_HD_CANVAS_SCALE;

  // ── Render each page ────────────────────────────────────────────────────
  const renderArea = document.getElementById('pdf-render-area');
  renderArea.innerHTML = '';

  const { jsPDF } = window.jspdf;
  const orientation = page.orientation === 'landscape' ? 'l' : 'p';
  const doc = new jsPDF({ orientation, unit: 'mm', format: [wMm, hMm], compress: false });

  for (let pi = 0; pi < totalPages; pi++) {
    const isFirst = pi === 0;
    const isLast  = pi === totalPages - 1;
    const slice   = pageSlices[pi];

    // Decide which elements go on this page
    const pageEls = [];

    // Header
    if (hEnabled && (hPages === 'all' || (hPages === 'first' && isFirst))) {
      pageEls.push(...headerEls);
    }

    // Footer
    if (fEnabled && (fPages === 'all' || (fPages === 'last' && isLast))) {
      pageEls.push(...footerEls);
    }

    // Static body (only page 1)
    if (isFirst) pageEls.push(...staticBodyEls);

    // Detail table
    let detailOverride = null;
    if (detailEl) {
        if (hasData && slice !== null) {
          const nextY = _detailStartYForPage(detailEl, page, pi);
          const repositioned = isFirst ? detailEl : { ...detailEl, y: nextY };
        pageEls.push(repositioned);
        detailOverride = { el: repositioned, rows: slice, allRows: detailRows };
      } else if (!hasData) {
        // No detail data — render table as designed
        pageEls.push(detailEl);
      }
    }

    // Build DOM
    const pageDOM = _buildPageDOM(page, wMm, hMm, pageEls, fieldValues, scale, detailOverride, pi + 1, totalPages);
    renderArea.innerHTML = '';
    renderArea.appendChild(pageDOM);

    // Wait for images
    await Promise.all(
      Array.from(pageDOM.querySelectorAll('img')).map(
        img => img.complete ? Promise.resolve()
          : new Promise(r => { img.onload = r; img.onerror = r; })
      )
    );

    // Capture
    let canvas;
    try {
      canvas = await html2canvas(pageDOM, {
        scale: canvasScale,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        logging: false,
        width: pageDOM.offsetWidth,
        height: pageDOM.offsetHeight,
      });
    } catch (err) {
      renderArea.innerHTML = '';
      throw new Error(`html2canvas failed on page ${pi + 1}: ${err.message}`);
    }

    const imgData = canvas.toDataURL('image/png');
    if (pi === 0) {
      doc.addImage(imgData, 'PNG', 0, 0, wMm, hMm, undefined, 'NONE');
    } else {
      doc.addPage([wMm, hMm], orientation);
      doc.addImage(imgData, 'PNG', 0, 0, wMm, hMm, undefined, 'NONE');
    }
  }

  renderArea.innerHTML = '';

  // Open preview tab first with explicit Download button using the desired filename.
  // If popup is blocked, fallback to direct download.
  const pdfBlob = doc.output('blob');
  const fileName = `${_sanitizePdfBaseName(layout?.name)}_${_pdfTimestampDDMMYYYYHHMMSS()}.pdf`;
  const blobUrl = URL.createObjectURL(pdfBlob);
  const opened = window.open('', '_blank');
  if (opened) {
    const safeName = _escapeHtml(fileName);
    opened.document.open();
    opened.document.write(`<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${safeName}</title>
  <style>
    body{margin:0;font-family:Arial,sans-serif;background:#f2f3f7;}
    .bar{height:44px;display:flex;align-items:center;justify-content:space-between;padding:0 12px;background:#ffffff;border-bottom:1px solid #d9dbe3;box-sizing:border-box;}
    .name{font-size:13px;color:#1f2937;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:70vw;}
    .btn{display:inline-flex;align-items:center;justify-content:center;height:30px;padding:0 12px;border:1px solid #2563eb;border-radius:6px;background:#2563eb;color:#fff;font-size:12px;text-decoration:none;}
    iframe{width:100%;height:calc(100vh - 44px);border:0;display:block;background:#fff;}
  </style>
</head>
<body>
  <div class="bar">
    <div class="name">${safeName}</div>
    <a class="btn" href="${blobUrl}" download="${safeName}">Download PDF</a>
  </div>
  <iframe src="${blobUrl}" title="${safeName}"></iframe>
</body>
</html>`);
    opened.document.close();
  } else {
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
  }
  setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
  return pdfBlob;
}

/**
 * Render a live preview of the layout into containerEl.
 * Builds the same page DOM as generatePDF but displays it directly (no html2canvas).
 */
async function renderLayoutPreview(layout, fieldValues, detailRows, containerEl) {
  const page = layout.page;
  const sizes = window.PAGE_SIZES;
  let wMm;
  let hMm;
  if (page.size === 'custom') {
    wMm = Math.max(20, parseFloat(page.customWidthMm ?? page.customWidth) || 210);
    hMm = Math.max(20, parseFloat(page.customHeightMm ?? page.customHeight) || 297);
  } else {
    ({ width: wMm, height: hMm } = sizes[page.size] || sizes['A4']);
  }
  if (page.orientation === 'landscape') [wMm, hMm] = [hMm, wMm];

  // Scale to fit container width with a little padding
  const containerWidth = Math.max(200, containerEl.clientWidth - 40);
  const scale = Math.min(PDF_MM_TO_PX, containerWidth / wMm);

  const allElements = layout.elements || [];
  const hEnabled = !!page.headerEnabled;
  const fEnabled = !!page.footerEnabled;
  const hH  = hEnabled ? (page.headerHeight || 20) : 0;
  const fH  = fEnabled ? (page.footerHeight || 15) : 0;
  const hPages = page.headerPages || 'all';
  const fPages = page.footerPages || 'all';

  const headerEls     = allElements.filter(el => _getElementZone(el, page, hMm) === 'header');
  const footerEls     = allElements.filter(el => _getElementZone(el, page, hMm) === 'footer');
  const bodyEls       = allElements.filter(el => _getElementZone(el, page, hMm) === 'body');
  const detailEl      = bodyEls.find(el => el.type === 'table' && el.table?.detailMode === true);
  const staticBodyEls = bodyEls.filter(el => el !== detailEl);
  const hasData       = detailEl && detailRows && detailRows.length > 0;

  let pageSlices;
  if (hasData) {
    pageSlices = _buildPageSlices(detailEl, fieldValues, scale, detailRows, page, hMm);
  } else {
    pageSlices = [null];
  }

  const totalPages = pageSlices.length;
  containerEl.innerHTML = '';

  for (let pi = 0; pi < totalPages; pi++) {
    const isFirst = pi === 0;
    const isLast  = pi === totalPages - 1;
    const slice   = pageSlices[pi];

    const pageEls = [];
    if (hEnabled && (hPages === 'all' || (hPages === 'first' && isFirst))) pageEls.push(...headerEls);
    if (fEnabled && (fPages === 'all' || (fPages === 'last'  && isLast)))  pageEls.push(...footerEls);
    if (isFirst) pageEls.push(...staticBodyEls);

    let detailOverride = null;
    if (detailEl) {
        if (hasData && slice !== null) {
          const nextY = _detailStartYForPage(detailEl, page, pi);
          const repositioned = isFirst ? detailEl : { ...detailEl, y: nextY };
        pageEls.push(repositioned);
        detailOverride = { el: repositioned, rows: slice, allRows: detailRows };
      } else if (!hasData) {
        pageEls.push(detailEl);
      }
    }

    if (totalPages > 1) {
      const lbl = document.createElement('div');
      lbl.className = 'preview-page-label';
      lbl.textContent = `Page ${pi + 1} of ${totalPages}`;
      containerEl.appendChild(lbl);
    }

    const pageDOM = _buildPageDOM(page, wMm, hMm, pageEls, fieldValues, scale, detailOverride, pi + 1, totalPages);

    // Zone guide overlays — match designer appearance so positions look identical
      if (_isHeaderActiveOnPage(page, pi)) {
      const hOv = document.createElement('div');
      hOv.style.cssText = `position:absolute;left:0;width:100%;top:0;height:${hH * scale}px;background:rgba(59,130,246,0.07);border-bottom:2px dashed rgba(59,130,246,0.5);box-sizing:border-box;pointer-events:none;z-index:10;`;
      const hLbl = document.createElement('span');
      hLbl.style.cssText = 'position:absolute;right:8px;bottom:2px;font-size:10px;font-weight:700;letter-spacing:0.08em;color:rgba(59,130,246,0.75);';
      hLbl.textContent = 'HEADER' + (page.headerPages === 'first' ? ' (1st page)' : '');
      hOv.appendChild(hLbl);
      pageDOM.appendChild(hOv);
    }
      if (_isFooterActiveOnPage(page, pi, totalPages)) {
      const fTop = (hMm - fH) * scale;
      const fOv = document.createElement('div');
      fOv.style.cssText = `position:absolute;left:0;width:100%;top:${fTop}px;height:${fH * scale}px;background:rgba(249,115,22,0.07);border-top:2px dashed rgba(249,115,22,0.5);box-sizing:border-box;pointer-events:none;z-index:10;`;
      const fLbl = document.createElement('span');
      fLbl.style.cssText = 'position:absolute;right:8px;top:2px;font-size:10px;font-weight:700;letter-spacing:0.08em;color:rgba(249,115,22,0.75);';
      fLbl.textContent = 'FOOTER' + (page.footerPages === 'last' ? ' (last page)' : '');
      fOv.appendChild(fLbl);
      pageDOM.appendChild(fOv);
    }

    const wrap = document.createElement('div');
    wrap.className = 'preview-page-wrap';
    wrap.style.width = (wMm * scale) + 'px';
    wrap.style.height = (hMm * scale) + 'px';
    wrap.appendChild(pageDOM);
    containerEl.appendChild(wrap);
  }

  // Wait for images
  await Promise.all(
    Array.from(containerEl.querySelectorAll('img')).map(img =>
      img.complete ? Promise.resolve() : new Promise(r => { img.onload = r; img.onerror = r; })
    )
  );
}

// Expose globally
window.generatePDF = generatePDF;
window.renderLayoutPreview = renderLayoutPreview;
window.applyFieldValues = applyFieldValues;
