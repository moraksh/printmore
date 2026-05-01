/**
 * pdf.js — PDF generation for PrintMore
 */

'use strict';

/**
 * pdf.js quick orientation:
 * - V2-only PDF pipeline.
 * - Canonical page planning is shared by Live Preview and PDF export.
 * - buildTableDOM() is the main visual table renderer used by both flows.
 *
 * Suggested reading order:
 * 1) _buildCanonicalRenderPlan()
 * 2) renderLayoutPreview()
 * 3) generatePDFV2() -> generatePDF()
 */

const PDF_MM_TO_PX = 3.7795; // same scale as designer for rendering
const PDF_BARCODE_CACHE_LIMIT = 600;

const _pdfBarcodeSvgCache = new Map();
const _pdfBarcodePngCache = new Map();

const _pdfTelemetry = [];
let _lastPdfTelemetry = null;

function _getPdfConfig() {
  return window.PRINTMORE_PDF_CONFIG || {};
}

function _resolvePdfProfile(layout) {
  const cfg = _getPdfConfig();
  const profiles = cfg.profiles || {};
  const fallback = cfg.defaults?.profile || 'standard';
  const raw = String(layout?.page?.pdfProfile || fallback).trim().toLowerCase();
  return profiles[raw] ? raw : fallback;
}

function _getPdfProfileConfig(profileId) {
  const cfg = _getPdfConfig();
  const profiles = cfg.profiles || {};
  return profiles[profileId] || profiles[cfg.defaults?.profile || 'standard'] || {
    id: 'standard',
    imageJpegQuality: 0.8,
    imageScale: 1.6,
    barcodeScale: 3,
    barcodeModuleWidth: 1.6,
    compress: true,
    maxBytesPerPage: 1.5 * 1024 * 1024,
    maxTotalBytes: 8 * 1024 * 1024,
    maxGenerationMs: 12000,
  };
}

function _normalizeRasterFormat(value, fallback = 'png') {
  const raw = String(value || fallback).trim().toLowerCase();
  return raw === 'jpeg' || raw === 'jpg' ? 'jpeg' : 'png';
}

function _getV2TableRasterOptions(profileCfg) {
  const scale = Math.max(1, Number(profileCfg?.tableRasterScale || profileCfg?.imageScale || 2));
  const format = _normalizeRasterFormat(profileCfg?.tableRasterFormat, 'png');
  const quality = Math.max(0.45, Math.min(0.98, Number(profileCfg?.tableRasterQuality || profileCfg?.imageJpegQuality || 0.82)));
  return { scale, format, quality };
}

function _getV2BarcodeRasterScale(profileCfg) {
  return Math.max(1, Number(profileCfg?.barcodeRasterScale || 2));
}

function _recordPdfTelemetry(entry) {
  const cfg = _getPdfConfig();
  const maxRecords = Math.max(10, parseInt(cfg?.telemetry?.maxInMemoryRecords, 10) || 100);
  const normalized = { timestamp: new Date().toISOString(), ...entry };
  _pdfTelemetry.push(normalized);
  while (_pdfTelemetry.length > maxRecords) _pdfTelemetry.shift();
  _lastPdfTelemetry = normalized;
  if (cfg?.telemetry?.verboseConsole) {
    try { console.log('[PrintMore][PDF]', normalized); } catch {}
  }
}

function _getLastPdfTelemetry() {
  return _lastPdfTelemetry ? { ..._lastPdfTelemetry } : null;
}

function _getPdfTelemetry() {
  return _pdfTelemetry.map(t => ({ ...t }));
}

function _clearPdfTelemetry() {
  _pdfTelemetry.length = 0;
  _lastPdfTelemetry = null;
}

function recordPdfEvent(entry) {
  _recordPdfTelemetry({ type: 'custom', ...(entry || {}) });
}

function evaluatePdfReleaseGate(records) {
  const cfg = _getPdfConfig();
  const gate = cfg.releaseGate || {};
  const list = Array.isArray(records) ? records : _pdfTelemetry;
  const pdfRuns = list.filter(r => r && r.type === 'pdf_generate' && r.engine === 'v2' && r.success !== false);
  const issues = [];
  const byProfile = {};
  pdfRuns.forEach(run => {
    const p = run.profile || 'standard';
    byProfile[p] = byProfile[p] || [];
    byProfile[p].push(run);
  });

  Object.entries(byProfile).forEach(([profile, runs]) => {
    const maxMs = Number(gate?.maxGenerationMs?.[profile] ?? Number.MAX_SAFE_INTEGER);
    const maxBytesPerPage = Number(gate?.maxBytesPerPage?.[profile] ?? Number.MAX_SAFE_INTEGER);
    runs.forEach(run => {
      if ((run.durationMs || 0) > maxMs) {
        issues.push(`Profile ${profile}: generation ${run.durationMs}ms exceeded ${maxMs}ms`);
      }
      if ((run.bytesPerPage || 0) > maxBytesPerPage) {
        issues.push(`Profile ${profile}: bytes/page ${run.bytesPerPage} exceeded ${Math.round(maxBytesPerPage)}`);
      }
    });
  });

  const overlapFailures = list.filter(r => r?.type === 'overlap_regression' && r?.failed).length;
  if (overlapFailures > (Number(gate.maxOverlapRegressions) || 0)) {
    issues.push(`Overlap regressions: ${overlapFailures} > ${gate.maxOverlapRegressions || 0}`);
  }

  const barcodeChecks = list.filter(r => r?.type === 'barcode_scan_check');
  if (barcodeChecks.length) {
    const okRate = barcodeChecks.filter(r => r?.ok).length / barcodeChecks.length;
    if (okRate < Number(gate.barcodeScanSuccessMin || 0)) {
      issues.push(`Barcode scan success ${(okRate * 100).toFixed(1)}% is below ${(Number(gate.barcodeScanSuccessMin || 0) * 100).toFixed(1)}%`);
    }
  }

  const visualChecks = list.filter(r => r?.type === 'visual_diff_check');
  if (visualChecks.length) {
    const maxDiff = Math.max(...visualChecks.map(v => Number(v.diffPct || 0)));
    if (maxDiff > Number(gate.visualDiffTolerancePct || Number.MAX_SAFE_INTEGER)) {
      issues.push(`Visual diff ${maxDiff}% exceeded tolerance ${gate.visualDiffTolerancePct}%`);
    }
  }

  return {
    ok: issues.length === 0,
    issues,
    summary: {
      totalRuns: pdfRuns.length,
      profiles: Object.keys(byProfile),
      overlapFailures,
      barcodeChecks: barcodeChecks.length,
      visualChecks: visualChecks.length,
    },
  };
}

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
    const val = _getFieldValueSmart(fieldValues, name);
    return (val !== undefined && val !== null && String(val) !== '') ? val : match;
  });
}

function _normalizeFieldName(name) {
  return String(name || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function _getFieldValueSmart(fieldValues, fieldName) {
  if (!fieldValues || !fieldName) return '';
  if (Object.prototype.hasOwnProperty.call(fieldValues, fieldName)) {
    return fieldValues[fieldName];
  }
  const wantedNorm = _normalizeFieldName(fieldName);
  for (const [k, v] of Object.entries(fieldValues)) {
    if (_normalizeFieldName(k) === wantedNorm) return v;
  }
  return '';
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
  const elTop = Number(el?.y) || 0;
  const elBottom = elTop + (Number(el?.height) || 0);
  // Use edge-based classification so tall elements keep the zone users intended
  // from their top placement in the designer.
  if (hH > 0 && elTop < hH) return 'header';
  if (fH > 0 && elBottom > footerStart) return 'footer';
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
  const hasPageBreakField = Boolean(String(page?.pageBreakField || '').trim());
  if (hasPageBreakField) {
    // In page-break mode, each page should keep the same designed body layout
    // (barcode/labels/table alignment) as page 1.
    return detailEl.y;
  }

  if (pageIndex === 0) return detailEl.y;

  const marginTop = page.marginTop ?? 15;
  const headerHeight = page.headerHeight || 0;
  const hasHeaderOnCurrentPage = _isHeaderActiveOnPage(page, pageIndex);
  // For continuation pages, start table rows at the current page body start only.
  // This avoids inheriting first-page body content spacing (customer/invoice/total block).
  return hasHeaderOnCurrentPage ? headerHeight : marginTop;
}

function _touchBarcodeCache(cache, key, value) {
  if (!cache || !key) return value;
  if (cache.has(key)) cache.delete(key);
  cache.set(key, value);
  while (cache.size > PDF_BARCODE_CACHE_LIMIT) {
    const oldest = cache.keys().next().value;
    if (!oldest) break;
    cache.delete(oldest);
  }
  return value;
}

function _barcodeCacheKey(value, opts) {
  return [
    String(value ?? ''),
    String(opts?.format || 'CODE128'),
    opts?.displayValue !== false ? '1' : '0',
    Number(opts?.fontSize || 0),
    String(opts?.lineColor || '#000000'),
    Number(opts?.width || 1),
    Number(opts?.height || 0),
    Number(opts?.margin || 0),
    Number(opts?.textMargin || 0),
    String(opts?.background || '#ffffff'),
  ].join('|');
}

function _getOrCreateBarcodeSvgMarkup(value, opts) {
  if (typeof JsBarcode === 'undefined') return null;
  const key = _barcodeCacheKey(value, opts);
  const cached = _pdfBarcodeSvgCache.get(key);
  if (cached) {
    _touchBarcodeCache(_pdfBarcodeSvgCache, key, cached);
    return cached;
  }
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  JsBarcode(svg, String(value || '0000000000'), opts || {});
  const svgW = parseFloat(svg.getAttribute('width')) || 200;
  const svgH = parseFloat(svg.getAttribute('height')) || 60;
  if (!svg.getAttribute('viewBox')) svg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);
  const markup = svg.outerHTML;
  return _touchBarcodeCache(_pdfBarcodeSvgCache, key, markup);
}

function _buildSvgNodeFromMarkup(markup) {
  if (!markup) return null;
  const holder = document.createElement('div');
  holder.innerHTML = markup;
  return holder.firstElementChild;
}

function _resolveLayoutPageSizeMm(layout) {
  const page = layout?.page || {};
  const sizes = window.PAGE_SIZES || {};
  let wMm;
  let hMm;
  if (page.size === 'custom') {
    wMm = Math.max(20, parseFloat(page.customWidthMm ?? page.customWidth) || 210);
    hMm = Math.max(20, parseFloat(page.customHeightMm ?? page.customHeight) || 297);
  } else {
    ({ width: wMm, height: hMm } = sizes[page.size] || sizes.A4 || { width: 210, height: 297 });
  }
  if (page.orientation === 'landscape') [wMm, hMm] = [hMm, wMm];
  return { wMm, hMm };
}

function _buildCanonicalRenderPlan(layout, fieldValues, detailRows, sliceScale = PDF_MM_TO_PX) {
  // This planner is the core parity layer:
  // both Live and PDF consume this exact page/element plan.
  const page = layout?.page || {};
  const { wMm, hMm } = _resolveLayoutPageSizeMm(layout);
  const allElements = layout?.elements || [];
  const hEnabled = !!page.headerEnabled;
  const fEnabled = !!page.footerEnabled;
  const mb = page.marginBottom ?? 15;
  const hPages = page.headerPages || 'all';
  const fPages = page.footerPages || 'all';

  const bodyEls = allElements.filter(el => _getElementZone(el, page, hMm) === 'body');
  const detailEl = bodyEls.find(el => el.type === 'table' && el.table?.detailMode === true);
  const hasData = !!(detailEl && Array.isArray(detailRows) && detailRows.length > 0);
  const hasPageBreakField = Boolean(String(page?.pageBreakField || '').trim());
  const pageSlices = hasData ? _buildPageSlices(detailEl, fieldValues, sliceScale, detailRows, page, hMm) : [null];
  if (hasData) checkDeterministicPageBreaks(detailEl, fieldValues, sliceScale, detailRows, page, hMm);
  const totalPages = pageSlices.length;

  const pages = [];
  for (let pi = 0; pi < totalPages; pi++) {
    const isFirst = pi === 0;
    const isLast = pi === totalPages - 1;
    const slice = pageSlices[pi];
    const footerActive = fEnabled && (fPages === 'all' || (fPages === 'last' && isLast));
    const detailBandTop = detailEl ? _detailStartYForPage(detailEl, page, pi) : 0;
    const detailBandBottom = detailEl ? (hMm - mb - (footerActive ? (page.footerHeight || 15) : 0)) : 0;
    const pageEls = [];
    let detailOverride = null;
    let pageFieldValues = fieldValues || {};

    // Single-page parity mode:
    // when there is no detail-driven pagination, render exactly what designer stored.
    if (!hasData && totalPages === 1) {
      pages.push({
        pageIndex: pi,
        pageNumber: pi + 1,
        totalPages,
        isFirst,
        isLast,
        footerActive,
        pageEls: allElements.slice(),
        fieldValues: pageFieldValues,
        detailOverride: null,
      });
      continue;
    }

    allElements.forEach(el => {
      const zone = _getElementZone(el, page, hMm);
      if (zone === 'header') {
        if (hEnabled && (hPages === 'all' || (hPages === 'first' && isFirst))) pageEls.push(el);
        return;
      }
      if (zone === 'footer') {
        if (fEnabled && (fPages === 'all' || (fPages === 'last' && isLast))) pageEls.push(el);
        return;
      }
      if (el === detailEl) {
        if (hasData && slice !== null) {
          const nextY = _detailStartYForPage(detailEl, page, pi);
          const repositioned = isFirst ? detailEl : { ...detailEl, y: nextY };
          pageEls.push(repositioned);
          detailOverride = { el: repositioned, rows: slice, allRows: detailRows };
          if (hasPageBreakField && Array.isArray(slice) && slice.length > 0) {
            // For page-break mode, non-table elements (barcode/fields/text placeholders)
            // should reflect the first row of that page slice.
            pageFieldValues = { ...(fieldValues || {}), ...(slice[0] || {}) };
          }
        } else if (!hasData) {
          pageEls.push(detailEl);
        }
        return;
      }

      if (!detailEl || !hasData) {
        pageEls.push(el);
      } else if (isFirst || hasPageBreakField) {
        // Keep non-detail body items on first page unless explicitly header/footer.
        // In page-break mode, repeat full body layout on each page/group.
        pageEls.push(el);
      }
    });

    pages.push({
      pageIndex: pi,
      pageNumber: pi + 1,
      totalPages,
      isFirst,
      isLast,
      footerActive,
      pageEls,
      fieldValues: pageFieldValues,
      detailOverride,
    });
  }

  return {
    page,
    wMm,
    hMm,
    pages,
    totalPages,
    detailEl,
    hasData,
  };
}

/**
 * Render all detail rows in a hidden container, measure each row's actual
 * DOM height, and return page slices with exactly the rows that fit.
 *
 * @returns {Array<Array>} slices — each entry is the detailRows for one page
 */
function _splitRowsByPageBreakField(detailRows, pageBreakField) {
  if (!Array.isArray(detailRows) || detailRows.length === 0) return [];
  const breakField = String(pageBreakField || '').trim();
  if (!breakField) return [detailRows];

  const groups = [];
  let current = [];
  let prevKey = null;
  detailRows.forEach((row, index) => {
    const raw = _getFieldValueSmart(row || {}, breakField);
    const key = String(raw ?? '').trim().toLowerCase();
    if (index === 0 || key === prevKey) {
      current.push(row);
    } else {
      groups.push(current);
      current = [row];
    }
    prevKey = key;
  });
  if (current.length) groups.push(current);
  return groups;
}

function _buildPageSlices(detailEl, fieldValues, scale, detailRows, page, hMm) {
  const mb = page.marginBottom ?? 15;
  const fH = page.footerEnabled ? (page.footerHeight || 0) : 0; // mm
  const footerPages = page.footerPages || 'all';
  const rowGroups = _splitRowsByPageBreakField(detailRows, page?.pageBreakField);

  // Render ALL rows into a hidden off-screen container to measure heights
  const tmp = document.createElement('div');
  tmp.style.cssText = `position:fixed;left:-9999px;top:0;width:${detailEl.width * scale}px;` +
    `visibility:hidden;pointer-events:none;overflow:visible;`;
  buildTableDOM(tmp, detailEl, fieldValues, scale, detailRows);
  document.body.appendChild(tmp);

  const tblEl = tmp.querySelector('table');
  const trs = tblEl ? Array.from(tblEl.querySelectorAll('tr')) : [];
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
  let globalRowIndex = 0; // index in original detailRows
  let pageIndex = 0;
  const groupsToProcess = rowGroups.length ? rowGroups : [detailRows];

  groupsToProcess.forEach(groupRows => {
    let localIndex = 0;
    while (localIndex < groupRows.length) {
      const avail = availablePx(pageIndex, null);
      let usedPx = 0;
      const batch = [];

      while (localIndex < groupRows.length) {
        const rh = rowHeightsPx[globalRowIndex + 1] || headerRowPx; // +1: trs[0] is header
        if (batch.length > 0 && usedPx + rh > avail + 1) break; // +1 rounding buffer
        batch.push(groupRows[localIndex]);
        usedPx += rh;
        localIndex++;
        globalRowIndex++;
      }

      // Guard: always advance at least 1 row to prevent infinite loop
      if (batch.length === 0 && localIndex < groupRows.length) {
        batch.push(groupRows[localIndex]);
        localIndex++;
        globalRowIndex++;
      }

      slices.push(batch);
      pageIndex++;
    }
  });

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

function checkDeterministicPageBreaks(detailEl, fieldValues, scale, detailRows, page, hMm) {
  if (!detailEl || !Array.isArray(detailRows)) {
    return { ok: true, fingerprintA: '', fingerprintB: '' };
  }
  const a = _buildPageSlices(detailEl, fieldValues, scale, detailRows, page, hMm).map(rows => rows.length).join('|');
  const b = _buildPageSlices(detailEl, fieldValues, scale, detailRows, page, hMm).map(rows => rows.length).join('|');
  const ok = a === b;
  _recordPdfTelemetry({
    type: 'page_break_determinism',
    ok,
    fingerprintA: a,
    fingerprintB: b,
    rows: detailRows.length,
  });
  return { ok, fingerprintA: a, fingerprintB: b };
}

function _measureTableHeightPx(el, fieldValues, scale, detailRows, allRows) {
  if (!el || el.type !== 'table') return 0;
  const tmp = document.createElement('div');
  tmp.style.cssText = `position:fixed;left:-9999px;top:0;width:${Math.max(1, el.width * scale)}px;` +
    `visibility:hidden;pointer-events:none;overflow:visible;`;
  buildTableDOM(tmp, el, fieldValues || {}, scale, detailRows || null, allRows || detailRows || null);
  document.body.appendChild(tmp);
  const tableEl = tmp.querySelector('table');
  const height = Math.ceil(tableEl?.getBoundingClientRect?.().height || 0);
  document.body.removeChild(tmp);
  return height;
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
  const detailEl = detailOverride?.el || null;
  const detailDesignedBottomPx = detailEl ? ((Number(detailEl.y) || 0) + (Number(detailEl.height) || 0)) * scale : null;
  const detailMeasuredHeightPx = detailEl
    ? _measureTableHeightPx(detailEl, fieldValues, scale, detailOverride?.rows || null, detailOverride?.allRows || detailOverride?.rows || null)
    : 0;
  const detailGrowthPx = detailEl ? Math.max(0, detailMeasuredHeightPx - ((Number(detailEl.height) || 0) * scale)) : 0;

  let pageBorderEl = null;
  if (page?.pageBorderEnabled) {
    const mt = Number(page.marginTop ?? 15);
    const mr = Number(page.marginRight ?? 15);
    const mb = Number(page.marginBottom ?? 15);
    const ml = Number(page.marginLeft ?? 15);
    const border = document.createElement('div');
    const ratio = scale / PDF_MM_TO_PX;
    const borderW = Math.max(1, Number(page.pageBorderWidth || 1) * ratio);
    const x = ml * scale;
    const y = mt * scale;
    const bw = Math.max(1, wPx - (ml + mr) * scale);
    const bh = Math.max(1, hPx - (mt + mb) * scale);
    border.style.cssText = `position:absolute;left:${x}px;top:${y}px;width:${bw}px;height:${bh}px;` +
      `box-sizing:border-box;border:${borderW}px solid #000000;pointer-events:none;z-index:999;`;
    pageBorderEl = border;
  }

  elements.forEach(el => {
    const isDetailEl = detailOverride && el === detailOverride.el;
    const wrapper = document.createElement('div');
    // el.x / el.y are in physical-page mm from the page top-left (same origin as the designer canvas).
    // Do NOT add margins here — the margins are already baked into the stored coordinates.
    const xPx = el.x * scale;
    let yPx = el.y * scale;
    // Flow stacked tables: when detail table expands, push later tables down.
    if (
      detailGrowthPx > 0 &&
      detailEl &&
      !isDetailEl &&
      el.type === 'table' &&
      yPx >= (detailDesignedBottomPx - 0.5)
    ) {
      yPx += detailGrowthPx;
    }
    const wElPx = el.width * scale;
    const hElPx = el.height * scale;
    wrapper.style.cssText = `position:absolute;left:${xPx}px;top:${yPx}px;width:${wElPx}px;height:${hElPx}px;box-sizing:border-box;overflow:hidden;opacity:${el.style?.opacity !== undefined ? el.style.opacity : 1};`;
    if (el.type === 'table') {
      wrapper.style.overflow = 'visible';
    }

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

  // Keep page border above all rendered content in live preview and PDF render DOM.
  if (pageBorderEl) container.appendChild(pageBorderEl);

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
  const ta = style.textAlign || 'left';
  domEl.style.textAlign = ta;
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
  if (ta === 'center') domEl.style.justifyContent = 'center';
  else if (ta === 'right') domEl.style.justifyContent = 'flex-end';
  else domEl.style.justifyContent = 'flex-start';
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
  const val = el.fieldName ? _getFieldValueSmart(fieldValues, el.fieldName) : '';
  inner.textContent = val;
  wrapper.appendChild(inner);
}

function buildUserDOM(wrapper, el) {
  const style = el.style || {};
  const inner = document.createElement('div');
  applyTextStyles(el, inner, style);
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

  wrapper.style.display = 'block';
  wrapper.style.position = 'absolute';
  wrapper.style.overflow = 'hidden';

  const line = document.createElement('div');
  if (direction === 'horizontal') {
    line.style.cssText = `position:absolute;left:0;top:0;width:100%;height:${thickness}px;background:${color};`;
  } else {
    line.style.cssText = `position:absolute;left:0;top:0;width:${thickness}px;height:100%;background:${color};`;
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

function _renderBarcodeCanvasToElement(targetEl, value, opts = {}) {
  if (typeof JsBarcode === 'undefined') return false;
  const widthPx = Math.max(16, Math.round(Number(opts.widthPx) || targetEl.clientWidth || 0));
  const heightPx = Math.max(12, Math.round(Number(opts.heightPx) || targetEl.clientHeight || 0));
  const showText = opts.showText !== false;
  const fontSize = Math.max(6, Number(opts.fontSize) || 10);
  const textAllowance = showText ? (fontSize + 5) : 0;
  const barHeight = Math.max(8, Math.round(heightPx - textAllowance - 4));

  const canvas = document.createElement('canvas');
  const scale = 2;
  canvas.width = Math.max(32, widthPx * scale);
  canvas.height = Math.max(24, heightPx * scale);
  const sample = String(value || '0000000000');
  const moduleWidth = Math.max(1, Math.floor((canvas.width - 8) / Math.max(24, sample.length * 11)));

  JsBarcode(canvas, sample, {
    format: opts.format || 'CODE128',
    displayValue: showText,
    fontSize: Math.max(8, Math.round(fontSize * scale)),
    lineColor: opts.lineColor || '#000000',
    height: Math.max(16, barHeight * scale),
    width: moduleWidth,
    margin: 2 * scale,
    textMargin: 2 * scale,
    background: '#ffffff',
  });

  canvas.style.cssText = 'display:block;width:100%;height:100%;';
  targetEl.style.overflow = 'hidden';
  targetEl.appendChild(canvas);
  return true;
}

function buildBarcodeDOM(wrapper, el, fieldValues) {
  const bc = el.barcode || {};
  const rawValue = el.fieldName
    ? String(_getFieldValueSmart(fieldValues, el.fieldName) || '')
    : (el.content || '');
  const value = rawValue || '0000000000';

  // Use individual property assignments — never reassign cssText (would reset position/size)
  wrapper.style.background = '#fff';
  wrapper.style.overflow = 'hidden';

  // Keep runtime rendering model aligned with designer with deterministic canvas sizing.
  try {
    const rendered = _renderBarcodeCanvasToElement(wrapper, value, {
      format: bc.type || 'CODE128',
      showText: bc.showText !== false,
      fontSize: bc.fontSize || 10,
      lineColor: bc.textColor || '#000000',
      widthPx: Math.max(16, Math.round(el.width * PDF_MM_TO_PX)),
      heightPx: Math.max(12, Math.round(el.height * PDF_MM_TO_PX)),
    });
    if (!rendered) throw new Error('barcode-render-failed');
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
  if (!isDetail) {
    tableEl.style.height = '100%';
    tableEl.style.boxSizing = 'border-box';
  }
  if (borderMode === 'all' || borderMode === 'outer') {
    tableEl.style.borderBottom = `${bw}px ${bs} ${bc}`;
  }

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
      td.style.boxSizing = 'border-box';
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
        // Outer frame
        if (r === 0) td.style.borderTop = `${bw}px ${bs} ${bc}`;
        if (r === rows - 1) td.style.borderBottom = `${bw}px ${bs} ${bc}`;
        if (c === 0) td.style.borderLeft = `${bw}px ${bs} ${bc}`;
        if (c === cols - 1 || (mergeNext && c + 1 === cols - 1)) td.style.borderRight = `${bw}px ${bs} ${bc}`;
        // Header divider
        if (r === 0) td.style.borderBottom = `2px ${bs} ${bc}`;
        // Footer divider (if footer row exists)
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
          if (dataCellDef?.fieldName && dataRows[r - 1]) cellValue = String(_getFieldValueSmart(dataRows[r - 1], dataCellDef.fieldName) || '');
          else if (dataCellDef?.content) cellValue = applyFieldValues(dataCellDef.content, fieldValues);
        } else {
          if (cellDef?.fieldName) cellValue = String(_getFieldValueSmart(fieldValues, cellDef.fieldName) || '');
          else if (cellDef?.content) cellValue = applyFieldValues(cellDef.content, fieldValues);
        }
        if (cellValue && typeof JsBarcode !== 'undefined') {
          td.style.padding = '2px';
          td.style.height = rowHpx + 'px';
          td.style.maxHeight = rowHpx + 'px';
          td.style.textAlign = 'center';
          try {
            const totalW = colWidths && colWidths.length === cols
              ? colWidths.reduce((s, w) => s + w, 0)
              : cols;
            const colRatio = colWidths && colWidths.length === cols
              ? (Number(colWidths[c] || 1) / Math.max(1, totalW))
              : (1 / Math.max(1, cols));
            const cellWidthPx = Math.max(16, Math.round((el.width * scale) * colRatio));
            const rendered = _renderBarcodeCanvasToElement(td, cellValue, {
              format: cp.barcodeType || 'CODE128',
              showText: cp.barcodeShowText !== false,
              fontSize: 7,
              lineColor: '#000000',
              widthPx: cellWidthPx,
              heightPx: rowHpx,
            });
            if (!rendered) throw new Error('barcode-render-failed');
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
          td.textContent = _getFieldValueSmart(dataRows[r - 1], dataCellDef.fieldName) || '';
        } else if (dataCellDef?.content) {
          td.textContent = applyFieldValues(dataCellDef.content, fieldValues);
        }
      } else {
        // Static mode
        if (cellDef?.fieldName) {
          td.textContent = _getFieldValueSmart(fieldValues, cellDef.fieldName) || '';
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
function _stringifyValue(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

function _resolveElementText(el, fieldValues, pageNum, totalPages) {
  if (!el) return '';
  if (el.type === 'text') return applyFieldValues(el.content || '', fieldValues);
  if (el.type === 'field') return _stringifyValue(el.fieldName ? fieldValues?.[el.fieldName] : '');
  if (el.type === 'user') return window.AuthStore?.currentUser?.()?.username || '';
  if (el.type === 'datetime') return _formatDateTimeStr(el.datetimeFormat || 'DD/MM/YYYY');
  if (el.type === 'pagenum') {
    const fmt = el.pagenumFormat || 'Page {n}';
    return fmt.replace('{n}', pageNum).replace('{total}', totalPages);
  }
  return '';
}

function _alignForText(style = {}) {
  const ta = String(style.textAlign || 'left').toLowerCase();
  if (ta === 'center') return { align: 'center', xFactor: 0.5 };
  if (ta === 'right') return { align: 'right', xFactor: 1 };
  return { align: 'left', xFactor: 0 };
}

function _normalizeFontFamilyName(rawFamily) {
  const raw = String(rawFamily || 'Arial').trim();
  const first = raw.split(',')[0].trim().replace(/^["']|["']$/g, '');
  return first.toLowerCase();
}

function _isPdfCoreFontFamily(fontFamily) {
  const ff = _normalizeFontFamilyName(fontFamily);
  return ff === 'helvetica' ||
    ff === 'arial' ||
    ff === 'times' ||
    ff === 'times new roman' ||
    ff === 'courier' ||
    ff === 'courier new';
}

function _wrapCanvasLines(ctx, text, maxWidthPx) {
  const words = String(text || '').split(/\s+/);
  if (!words.length) return [''];
  const lines = [];
  let line = words[0] || '';
  for (let i = 1; i < words.length; i++) {
    const test = `${line} ${words[i]}`;
    if (ctx.measureText(test).width <= maxWidthPx) {
      line = test;
    } else {
      lines.push(line);
      line = words[i];
    }
  }
  lines.push(line);
  return lines;
}

function _drawRasterizedTextOnly(doc, text, x, y, w, h, style = {}, options = {}) {
  const scale = 2;
  const wPx = Math.max(8, Math.round(w * PDF_MM_TO_PX * scale));
  const hPx = Math.max(8, Math.round(h * PDF_MM_TO_PX * scale));
  const canvas = document.createElement('canvas');
  canvas.width = wPx;
  canvas.height = hPx;
  const ctx = canvas.getContext('2d');
  if (!ctx) return false;

  ctx.clearRect(0, 0, wPx, hPx);
  const fontPt = Math.max(5, Number(style.fontSize || 10) || 10);
  const fontPx = fontPt * (96 / 72) * scale;
  const weight = String(style.fontWeight || 'normal').toLowerCase() === 'bold' ? '700' : '400';
  const italic = String(style.fontStyle || 'normal').toLowerCase() === 'italic' ? 'italic ' : '';
  const family = String(style.fontFamily || 'Arial');
  ctx.font = `${italic}${weight} ${fontPx}px ${family}`;
  const [r, g, b] = _hexToRgb(style.color || '#000000');
  ctx.fillStyle = `rgb(${r},${g},${b})`;
  ctx.textBaseline = 'top';

  const align = _alignForText(style);
  const pad = Math.max(1, Math.round(1.2 * scale));
  const maxTextWidth = Math.max(1, wPx - pad * 2);
  const lines = _wrapCanvasLines(ctx, _stringifyValue(text), maxTextWidth);
  const lineHeight = Math.round(fontPx * 1.25);
  let startY = pad;
  if (options.verticalAlign === 'middle') {
    startY = Math.max(pad, Math.round((hPx - lines.length * lineHeight) / 2));
  }

  lines.forEach((line, i) => {
    const yPx = startY + i * lineHeight;
    if (yPx + lineHeight > hPx) return;
    let xPx = pad;
    if (align.align === 'center') xPx = wPx / 2;
    if (align.align === 'right') xPx = wPx - pad;
    ctx.textAlign = align.align;
    ctx.fillText(line, xPx, yPx);
    const deco = String(style.textDecoration || 'none').toLowerCase();
    if (deco.includes('underline')) {
      ctx.strokeStyle = ctx.fillStyle;
      ctx.lineWidth = Math.max(1, Math.round(scale * 0.7));
      ctx.beginPath();
      const lw = ctx.measureText(line).width;
      const lx = align.align === 'left' ? xPx : (align.align === 'center' ? xPx - lw / 2 : xPx - lw);
      ctx.moveTo(lx, yPx + fontPx + 1);
      ctx.lineTo(lx + lw, yPx + fontPx + 1);
      ctx.stroke();
    }
  });

  doc.addImage(canvas.toDataURL('image/png'), 'PNG', x, y, w, h, undefined, 'FAST');
  return true;
}

function _drawV2TextElement(doc, el, text) {
  const s = el.style || {};
  const x = Number(el.x) || 0;
  const y = Number(el.y) || 0;
  const w = Math.max(0.1, Number(el.width) || 0);
  const h = Math.max(0.1, Number(el.height) || 0);

  if (s.backgroundColor && s.backgroundColor !== 'transparent') {
    _withRgb(doc, s.backgroundColor, (r, g, b) => {
      doc.setFillColor(r, g, b);
      doc.rect(x, y, w, h, 'F');
    }, [255, 255, 255]);
  }
  if ((s.borderWidth || 0) > 0) {
    _withRgb(doc, s.borderColor || '#000000', (r, g, b) => {
      doc.setDrawColor(r, g, b);
      doc.setLineWidth(_strokeWidthMmFromStylePx(s.borderWidth, 0.1, 0.7));
      _applyPdfBorderStyle(doc, s.borderStyle || 'solid');
      doc.rect(x, y, w, h);
      doc.setLineDashPattern([], 0);
    });
  }

  // For non-core fonts (e.g., Impact), rasterize text from browser font stack to preserve fidelity.
  if (!_isPdfCoreFontFamily(s.fontFamily)) {
    _drawRasterizedTextOnly(doc, _stringifyValue(text), x, y, w, h, s, { verticalAlign: 'top' });
    return;
  }

  _setPdfFont(doc, s, 10);
  const pad = 0.3;
  const textValue = _stringifyValue(text);
  const wrap = doc.splitTextToSize(textValue, Math.max(0.1, w - pad * 2));
  const fs = Math.max(5, Number(s.fontSize || 10) || 10);
  const lineHeight = fs * 0.36;
  // Keep text anchored near the top of its box for consistent layout parity.
  // so small top-right metadata does not collide with nearby lines.
  const startY = y + pad;
  const { align, xFactor } = _alignForText(s);
  const tx = x + pad + (w - pad * 2) * xFactor;
  wrap.forEach((line, i) => {
    const ly = startY + i * lineHeight;
    if (ly > y + h - 0.2) return;
    doc.text(line, tx, ly, { align, baseline: 'top', maxWidth: Math.max(0.1, w - pad * 2) });
    _applyTextDecorations(doc, line, x + pad, ly + fs * 0.05, Math.max(0.1, w - pad * 2), s);
  });
}

function _drawV2RectElement(doc, el) {
  const s = el.style || {};
  const x = Number(el.x) || 0;
  const y = Number(el.y) || 0;
  const w = Math.max(0.1, Number(el.width) || 0);
  const h = Math.max(0.1, Number(el.height) || 0);
  if (s.backgroundColor && s.backgroundColor !== 'transparent') {
    _withRgb(doc, s.backgroundColor, (r, g, b) => {
      doc.setFillColor(r, g, b);
      doc.rect(x, y, w, h, 'F');
    }, [255, 255, 255]);
  }
  if ((s.borderWidth || 0) > 0) {
    _withRgb(doc, s.borderColor || '#000000', (r, g, b) => {
      doc.setDrawColor(r, g, b);
      doc.setLineWidth(_strokeWidthMmFromStylePx(s.borderWidth, 0.1, 0.7));
      _applyPdfBorderStyle(doc, s.borderStyle || 'solid');
      doc.rect(x, y, w, h);
      doc.setLineDashPattern([], 0);
    });
  }
}

function _drawV2LineElement(doc, el) {
  const s = el.style || {};
  const x = Number(el.x) || 0;
  const y = Number(el.y) || 0;
  const w = Math.max(0.1, Number(el.width) || 0);
  const h = Math.max(0.1, Number(el.height) || 0);
  const direction = el.lineDirection || 'horizontal';
  _withRgb(doc, s.borderColor || '#000000', (r, g, b) => {
    doc.setDrawColor(r, g, b);
    doc.setLineWidth(_strokeWidthMmFromStylePx(s.borderWidth, 0.1, 0.7));
    _applyPdfBorderStyle(doc, s.borderStyle || 'solid');
  });
  if (direction === 'vertical') doc.line(x, y, x, y + h);
  else doc.line(x, y, x + w, y);
  doc.setLineDashPattern([], 0);
}

async function _rasterizeImageDataForV2(imageData, profileCfg) {
  if (!imageData) return null;
  if (typeof imageData !== 'string') return null;
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.max(1, Number(profileCfg?.imageScale || 1.6));
      const cw = Math.max(1, Math.round(img.width * scale));
      const ch = Math.max(1, Math.round(img.height * scale));
      const c = document.createElement('canvas');
      c.width = cw;
      c.height = ch;
      const ctx = c.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, c.width, c.height);
      ctx.drawImage(img, 0, 0, c.width, c.height);
      resolve({
        dataUrl: c.toDataURL('image/jpeg', Math.max(0.55, Math.min(0.95, Number(profileCfg?.imageJpegQuality || 0.8)))),
        format: 'JPEG',
        srcWidth: c.width,
        srcHeight: c.height,
      });
    };
    img.onerror = () => resolve({
      dataUrl: imageData,
      format: imageData.startsWith('data:image/png') ? 'PNG' : 'JPEG',
      srcWidth: null,
      srcHeight: null,
    });
    img.src = imageData;
  });
}

async function _rasterizeDomElementForV2(widthMm, heightMm, buildFn, options = {}) {
  const scale = Number(options.scale || 2);
  const opacity = options.opacity;
  const outputType = _normalizeRasterFormat(options.imageType, 'png');
  const outputQuality = Math.max(0.45, Math.min(0.98, Number(options.imageQuality || 0.82)));
  const wPx = Math.max(1, Math.round(widthMm * PDF_MM_TO_PX));
  const hPx = Math.max(1, Math.round(heightMm * PDF_MM_TO_PX));

  const host = document.createElement('div');
  host.style.cssText = `position:fixed;left:-9999px;top:0;width:${wPx}px;height:${hPx}px;` +
    `overflow:hidden;pointer-events:none;visibility:visible;box-sizing:border-box;background:#ffffff;`;
  if (Number.isFinite(opacity)) host.style.opacity = String(opacity);
  document.body.appendChild(host);

  try {
    await buildFn(host, PDF_MM_TO_PX);
    await Promise.all(
      Array.from(host.querySelectorAll('img')).map(img =>
        img.complete ? Promise.resolve() : new Promise(r => { img.onload = r; img.onerror = r; })
      )
    );
    const canvas = await html2canvas(host, {
      scale,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#ffffff',
      logging: false,
      width: wPx,
      height: hPx,
    });
    const dataUrl = outputType === 'jpeg'
      ? canvas.toDataURL('image/jpeg', outputQuality)
      : canvas.toDataURL('image/png');
    return {
      dataUrl,
      format: outputType === 'jpeg' ? 'JPEG' : 'PNG',
      widthPx: canvas.width || wPx,
      heightPx: canvas.height || hPx,
    };
  } finally {
    document.body.removeChild(host);
  }
}

async function _drawV2ImageElement(doc, el, profileCfg) {
  const img = await _rasterizeImageDataForV2(el.imageData, profileCfg);
  if (!img?.dataUrl) return;
  const s = el.style || {};
  const x = Number(el.x) || 0;
  const y = Number(el.y) || 0;
  const w = Math.max(0.1, Number(el.width) || 0);
  const h = Math.max(0.1, Number(el.height) || 0);

  if (s.backgroundColor && s.backgroundColor !== 'transparent') {
    _withRgb(doc, s.backgroundColor, (r, g, b) => {
      doc.setFillColor(r, g, b);
      doc.rect(x, y, w, h, 'F');
    }, [255, 255, 255]);
  }

  // Match designer's object-fit: contain behavior.
  let drawW = w;
  let drawH = h;
  if (img.srcWidth && img.srcHeight) {
    const srcRatio = img.srcWidth / img.srcHeight;
    const dstRatio = w / h;
    if (srcRatio > dstRatio) {
      drawW = w;
      drawH = w / srcRatio;
    } else {
      drawH = h;
      drawW = h * srcRatio;
    }
  }
  const dx = x + (w - drawW) / 2;
  const dy = y + (h - drawH) / 2;
  doc.addImage(img.dataUrl, img.format || 'JPEG', dx, dy, drawW, drawH, undefined, 'FAST');

  if ((s.borderWidth || 0) > 0) {
    _withRgb(doc, s.borderColor || '#000000', (r, g, b) => {
      doc.setDrawColor(r, g, b);
      doc.setLineWidth(_strokeWidthMmFromStylePx(s.borderWidth, 0.1, 0.7));
      _applyPdfBorderStyle(doc, s.borderStyle || 'solid');
      doc.rect(x, y, w, h);
      doc.setLineDashPattern([], 0);
    });
  }
}

function _buildBarcodeDataForV2(el, value, bc, profileCfg) {
  const scale = Math.max(1, Number(profileCfg?.barcodeScale || 3) / 2);
  const hPx = Math.max(8, (Number(el?.height) || 20) * PDF_MM_TO_PX * scale);
  const fontSize = (Number(bc?.fontSize) || 10) * scale;
  const textAllowance = (bc?.showText !== false) ? (fontSize + 4 * scale) : 0;
  const barHeight = Math.max(8 * scale, Math.round(hPx - textAllowance - 6 * scale));
  const opts = {
    format: bc?.type || 'CODE128',
    displayValue: bc?.showText !== false,
    lineColor: bc?.textColor || '#000000',
    background: '#ffffff',
    width: 1.5 * scale,
    height: barHeight,
    margin: 2 * scale,
    fontSize,
    textMargin: 2 * scale,
  };
  const key = _barcodeCacheKey(value, opts);
  const cached = _pdfBarcodePngCache.get(key);
  if (cached) {
    _touchBarcodeCache(_pdfBarcodePngCache, key, cached);
    return cached;
  }
  const canvas = document.createElement('canvas');
  JsBarcode(canvas, String(value || '0000000000'), opts);
  const payload = {
    dataUrl: canvas.toDataURL('image/png'),
    width: canvas.width || 1,
    height: canvas.height || 1,
  };
  return _touchBarcodeCache(_pdfBarcodePngCache, key, payload);
}

function _drawContainedPngInBox(doc, pngDataUrl, srcW, srcH, x, y, w, h, padding = 0.4) {
  const bx = x + padding;
  const by = y + padding;
  const bw = Math.max(0.1, w - padding * 2);
  const bh = Math.max(0.1, h - padding * 2);
  const safeW = Math.max(1, Number(srcW) || 1);
  const safeH = Math.max(1, Number(srcH) || 1);
  const srcRatio = safeW / safeH;
  const boxRatio = bw / bh;
  let drawW = bw;
  let drawH = bh;
  if (srcRatio > boxRatio) {
    drawW = bw;
    drawH = bw / srcRatio;
  } else {
    drawH = bh;
    drawW = bh * srcRatio;
  }
  const dx = bx + (bw - drawW) / 2;
  const dy = by + (bh - drawH) / 2;
  doc.addImage(pngDataUrl, 'PNG', dx, dy, drawW, drawH, undefined, 'FAST');
}

async function _drawV2BarcodeElement(doc, el, fieldValues, profileCfg) {
  if (typeof JsBarcode === 'undefined') {
    _drawV2TextElement(doc, el, _resolveElementText({ ...el, type: 'field' }, fieldValues));
    return;
  }
  const value = el.fieldName
    ? _stringifyValue(fieldValues?.[el.fieldName])
    : _stringifyValue(el.content || '');
  const bc = el.barcode || {};
  const localEl = {
    ...el,
    content: el.fieldName ? '' : (value || ''),
    fieldName: el.fieldName || '',
    barcode: {
      ...bc,
      type: bc.type || 'CODE128',
      showText: bc.showText !== false,
      fontSize: bc.fontSize || 10,
      textColor: bc.textColor || '#000000',
    },
  };
  try {
    const raster = await _rasterizeDomElementForV2(
      Math.max(0.1, Number(el.width) || 0),
      Math.max(0.1, Number(el.height) || 0),
      (host) => {
        host.style.position = 'relative';
        buildBarcodeDOM(host, localEl, fieldValues || {});
      },
      {
        scale: _getV2BarcodeRasterScale(profileCfg),
        opacity: localEl.style?.opacity,
        imageType: 'png',
      }
    );
    doc.addImage(
      raster.dataUrl,
      raster.format || 'PNG',
      Number(el.x) || 0,
      Number(el.y) || 0,
      Math.max(0.1, Number(el.width) || 0),
      Math.max(0.1, Number(el.height) || 0),
      undefined,
      'FAST'
    );
  } catch {
    const data = _buildBarcodeDataForV2(el, value || '0000000000', bc, profileCfg);
    _drawContainedPngInBox(
      doc,
      data.dataUrl,
      data.width,
      data.height,
      Number(el.x) || 0,
      Number(el.y) || 0,
      Math.max(0.1, Number(el.width) || 0),
      Math.max(0.1, Number(el.height) || 0),
      0.4
    );
  }
}

async function _drawV2TableElement(doc, el, fieldValues, detailRows, allDetailRows, profileCfg) {
  const x = Number(el.x) || 0;
  const y = Number(el.y) || 0;
  const w = Math.max(0.1, Number(el.width) || 0);
  const h = Math.max(0.1, Number(el.height) || 0);
  const isDetail = el?.table?.detailMode === true;
  const tableRaster = _getV2TableRasterOptions(profileCfg);
  try {
    // Detail tables can grow beyond designed table box for repeating rows.
    // Rasterize with dynamic content height so rows are not clipped in PDF pages.
    if (isDetail && Array.isArray(detailRows) && detailRows.length) {
      const scale = tableRaster.scale;
      const wPx = Math.max(1, Math.round(w * PDF_MM_TO_PX));
      const fallbackHPx = Math.max(1, Math.round(h * PDF_MM_TO_PX));
      const host = document.createElement('div');
      host.style.cssText = `position:fixed;left:-9999px;top:0;width:${wPx}px;` +
        `overflow:visible;pointer-events:none;visibility:visible;box-sizing:border-box;background:#ffffff;`;
      if (Number.isFinite(el?.style?.opacity)) host.style.opacity = String(el.style.opacity);
      document.body.appendChild(host);
      try {
        host.style.position = 'relative';
        const inner = document.createElement('div');
        inner.style.cssText = `position:relative;width:${wPx}px;height:auto;overflow:visible;box-sizing:border-box;`;
        host.appendChild(inner);
        buildTableDOM(inner, el, fieldValues || {}, PDF_MM_TO_PX, detailRows || null, allDetailRows || detailRows || null);
        await Promise.all(
          Array.from(host.querySelectorAll('img')).map(img =>
            img.complete ? Promise.resolve() : new Promise(r => { img.onload = r; img.onerror = r; })
          )
        );

        const tableEl = host.querySelector('table');
        const contentHPx = Math.max(
          fallbackHPx,
          Math.ceil(tableEl?.getBoundingClientRect()?.height || 0),
          Math.ceil(inner.getBoundingClientRect()?.height || 0),
          Math.ceil(inner.scrollHeight || 0)
        );

        host.style.height = contentHPx + 'px';
        inner.style.height = contentHPx + 'px';

        const canvas = await html2canvas(host, {
          scale,
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          logging: false,
          width: wPx,
          height: contentHPx,
        });
        const drawHmm = Math.max(0.1, contentHPx / PDF_MM_TO_PX);
        const imgData = tableRaster.format === 'jpeg'
          ? canvas.toDataURL('image/jpeg', tableRaster.quality)
          : canvas.toDataURL('image/png');
        doc.addImage(
          imgData,
          tableRaster.format === 'jpeg' ? 'JPEG' : 'PNG',
          x,
          y,
          w,
          drawHmm,
          undefined,
          'FAST'
        );
      } finally {
        document.body.removeChild(host);
      }
      return;
    }

    const raster = await _rasterizeDomElementForV2(
      w,
      h,
      (host) => {
        host.style.position = 'relative';
        const inner = document.createElement('div');
        inner.style.cssText = `position:relative;width:${Math.max(1, Math.round(w * PDF_MM_TO_PX))}px;height:${Math.max(1, Math.round(h * PDF_MM_TO_PX))}px;overflow:hidden;box-sizing:border-box;`;
        host.appendChild(inner);
        buildTableDOM(inner, el, fieldValues || {}, PDF_MM_TO_PX, detailRows || null, allDetailRows || detailRows || null);
      },
      {
        scale: tableRaster.scale,
        opacity: el?.style?.opacity,
        imageType: tableRaster.format,
        imageQuality: tableRaster.quality,
      }
    );
    doc.addImage(raster.dataUrl, raster.format || 'PNG', x, y, w, h, undefined, 'FAST');
    return;
  } catch {
    // Fail-soft: keep V2 resilient by drawing fallback border if rasterization fails.
    const style = el.style || {};
    _withRgb(doc, style.borderColor || '#000000', (r, g, b) => {
      doc.setDrawColor(r, g, b);
      doc.setLineWidth(_strokeWidthMmFromStylePx(style.borderWidth || 1, 0.1, 0.7));
      _applyPdfBorderStyle(doc, style.borderStyle || 'solid');
      doc.rect(x, y, w, h);
      doc.setLineDashPattern([], 0);
    });
  }
}

function _openPdfPreview(doc, layout) {
  const pdfBlob = doc.output('blob');
  const fileName = `${_sanitizePdfBaseName(layout?.name)}_${_pdfTimestampDDMMYYYYHHMMSS()}.pdf`;
  const blobUrl = URL.createObjectURL(pdfBlob);
  const opened = window.open('', '_blank');
  if (opened) {
    const safeName = _escapeHtml(fileName);
    const jsFileName = JSON.stringify(fileName);
    const jsBlobUrl = JSON.stringify(blobUrl);
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
    .btn{display:inline-flex;align-items:center;justify-content:center;height:30px;padding:0 12px;border:1px solid #2563eb;border-radius:6px;background:#2563eb;color:#fff;font-size:12px;text-decoration:none;cursor:pointer;}
    iframe{width:100%;height:calc(100vh - 44px);border:0;display:block;background:#fff;}
  </style>
</head>
<body>
  <div class="bar">
    <div class="name">${safeName}</div>
    <button class="btn" id="downloadBtn" type="button">Download PDF</button>
  </div>
  <iframe src="${blobUrl}" title="${safeName}"></iframe>
  <script>
    (function () {
      var fileName = ${jsFileName};
      var blobUrl = ${jsBlobUrl};
      var btn = document.getElementById('downloadBtn');
      if (!btn) return;
      btn.addEventListener('click', async function () {
        try {
          var resp = await fetch(blobUrl);
          if (!resp.ok) throw new Error('Download source not available');
          var blob = await resp.blob();
          var localUrl = URL.createObjectURL(blob);
          var a = document.createElement('a');
          a.href = localUrl;
          a.download = fileName;
          document.body.appendChild(a);
          a.click();
          a.remove();
          setTimeout(function () { URL.revokeObjectURL(localUrl); }, 5000);
        } catch (e) {
          alert('Could not download PDF. Please use browser save/download icon.');
        }
      });
    })();
  </script>
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

async function generatePDFV2(layout, fieldValues, detailRows, options = {}) {
  const startedAt = performance.now();
  const profileId = options.profile || _resolvePdfProfile(layout);
  const profileCfg = _getPdfProfileConfig(profileId);
  const plan = _buildCanonicalRenderPlan(layout, fieldValues, detailRows, PDF_MM_TO_PX);
  const { page, wMm, hMm, pages, totalPages } = plan;
  const orientation = page.orientation === 'landscape' ? 'l' : 'p';
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation, unit: 'mm', format: [wMm, hMm], compress: profileCfg.compress !== false });
  // Render each canonical page to image once, then place into jsPDF page.
  // This keeps PDF output aligned with Live preview layout behavior.
  const pageRaster = _getV2TableRasterOptions(profileCfg);
  const wPx = Math.max(1, Math.round(wMm * PDF_MM_TO_PX));
  const hPx = Math.max(1, Math.round(hMm * PDF_MM_TO_PX));

  for (let pi = 0; pi < totalPages; pi++) {
    if (pi > 0) doc.addPage([wMm, hMm], orientation);
    const pagePlan = pages[pi];
    const pageDOM = _buildPageDOM(
      page,
      wMm,
      hMm,
      pagePlan.pageEls,
      pagePlan.fieldValues || fieldValues,
      PDF_MM_TO_PX,
      pagePlan.detailOverride,
      pagePlan.pageNumber,
      pagePlan.totalPages
    );
    const renderHost = document.createElement('div');
    renderHost.style.cssText = `position:fixed;left:-9999px;top:0;width:${wPx}px;height:${hPx}px;overflow:hidden;background:#ffffff;`;
    renderHost.appendChild(pageDOM);
    document.body.appendChild(renderHost);
    try {
      await Promise.all(
        Array.from(renderHost.querySelectorAll('img')).map(img =>
          img.complete ? Promise.resolve() : new Promise(r => { img.onload = r; img.onerror = r; })
        )
      );
      const canvas = await html2canvas(renderHost, {
        scale: pageRaster.scale,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        logging: false,
        width: wPx,
        height: hPx,
      });
      const imgData = pageRaster.format === 'jpeg'
        ? canvas.toDataURL('image/jpeg', pageRaster.quality)
        : canvas.toDataURL('image/png');
      doc.addImage(
        imgData,
        pageRaster.format === 'jpeg' ? 'JPEG' : 'PNG',
        0,
        0,
        wMm,
        hMm,
        undefined,
        'FAST'
      );
    } finally {
      document.body.removeChild(renderHost);
    }
  }

  const pdfBlob = _openPdfPreview(doc, layout);
  const durationMs = Math.round(performance.now() - startedAt);
  const bytes = pdfBlob.size || 0;
  const bytesPerPage = totalPages > 0 ? Math.round(bytes / totalPages) : bytes;
  const maxTotal = Number(profileCfg.maxTotalBytes || 0) || Number.MAX_SAFE_INTEGER;

  _recordPdfTelemetry({
    type: 'pdf_generate',
    engine: 'v2',
    profile: profileId,
    layoutId: layout?.id || '',
    layoutName: layout?.name || '',
    pages: totalPages,
    detailRows: Array.isArray(detailRows) ? detailRows.length : 0,
    bytes,
    bytesPerPage,
    durationMs,
    success: true,
  });

  return {
    blob: pdfBlob,
    meta: {
      engine: 'v2',
      profile: profileId,
      bytes,
      pages: totalPages,
      bytesPerPage,
      durationMs,
      exceededSizeGate: bytes > maxTotal,
    },
  };
}

function _hexToRgb(color, fallback = [0, 0, 0]) {
  if (!color || typeof color !== 'string') return fallback;
  const c = color.trim();
  if (/^#([0-9a-f]{3})$/i.test(c)) {
    const m = c.slice(1);
    return [
      parseInt(m[0] + m[0], 16),
      parseInt(m[1] + m[1], 16),
      parseInt(m[2] + m[2], 16),
    ];
  }
  if (/^#([0-9a-f]{6})$/i.test(c)) {
    return [
      parseInt(c.slice(1, 3), 16),
      parseInt(c.slice(3, 5), 16),
      parseInt(c.slice(5, 7), 16),
    ];
  }
  const rgb = c.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
  if (rgb) {
    return [
      Math.max(0, Math.min(255, parseInt(rgb[1], 10) || 0)),
      Math.max(0, Math.min(255, parseInt(rgb[2], 10) || 0)),
      Math.max(0, Math.min(255, parseInt(rgb[3], 10) || 0)),
    ];
  }
  return fallback;
}

function _withRgb(doc, color, fn, fallback = [0, 0, 0]) {
  const [r, g, b] = _hexToRgb(color, fallback);
  fn(r, g, b);
}

function _setPdfFont(doc, style = {}, fallbackSize = 10) {
  const ffRaw = String(style.fontFamily || 'Helvetica').toLowerCase();
  let family = 'helvetica';
  if (ffRaw.includes('times') || ffRaw.includes('georgia') || ffRaw.includes('merriweather') || ffRaw.includes('cambria')) family = 'times';
  if (ffRaw.includes('courier') || ffRaw.includes('mono')) family = 'courier';
  const bold = String(style.fontWeight || 'normal').toLowerCase() === 'bold';
  const italic = String(style.fontStyle || 'normal').toLowerCase() === 'italic';
  const fontStyle = bold && italic ? 'bolditalic' : bold ? 'bold' : italic ? 'italic' : 'normal';
  doc.setFont(family, fontStyle);
  doc.setFontSize(Math.max(5, Number(style.fontSize || fallbackSize) || fallbackSize));
  _withRgb(doc, style.color || '#000000', (r, g, b) => doc.setTextColor(r, g, b));
}

function _strokeWidthMmFromStylePx(widthPx, minMm = 0.1, maxMm = 0.8) {
  const px = Number(widthPx);
  const safePx = Number.isFinite(px) ? px : 1;
  const mm = safePx * 0.264583; // CSS px -> mm
  return Math.max(minMm, Math.min(maxMm, mm));
}

function _applyPdfBorderStyle(doc, borderStyle) {
  const bs = String(borderStyle || 'solid').toLowerCase();
  if (bs === 'dashed') {
    doc.setLineDashPattern([1.2, 0.8], 0);
  } else if (bs === 'dotted') {
    doc.setLineDashPattern([0.25, 0.7], 0);
  } else {
    doc.setLineDashPattern([], 0);
  }
}

function _applyTextDecorations(doc, text, x, y, maxWidthMm, style) {
  const deco = String(style?.textDecoration || 'none').toLowerCase();
  if (deco === 'none') return;
  const fs = Math.max(5, Number(style?.fontSize || 10) || 10);
  const textWidth = Math.min(doc.getTextWidth(text), maxWidthMm);
  const weight = _strokeWidthMmFromStylePx(style?.borderWidth || 0.3, 0.1, 0.5);
  _applyPdfBorderStyle(doc, style?.borderStyle || 'solid');
  if (deco.includes('underline')) {
    doc.setLineWidth(weight);
    doc.line(x, y + (fs * 0.12), x + textWidth, y + (fs * 0.12));
  }
  if (deco.includes('line-through')) {
    doc.setLineWidth(weight);
    doc.line(x, y - (fs * 0.28), x + textWidth, y - (fs * 0.28));
  }
  doc.setLineDashPattern([], 0);
}

/**
 * Public PDF generation API.
 * V2-only renderer.
 */
async function generatePDF(layout, fieldValues, detailRows) {
  const startedAt = performance.now();
  const requestedProfile = _resolvePdfProfile(layout);
  try {
    const out = await generatePDFV2(layout, fieldValues, detailRows, { profile: requestedProfile });
    if (!out?.blob) throw new Error('V2 PDF did not return a valid blob.');
    const meta = out.meta || {};
    const profileCfg = _getPdfProfileConfig(meta.profile || requestedProfile);
    if (meta.exceededSizeGate && requestedProfile === 'print_hd') {
      if (window.showToast) window.showToast('Print HD exceeded size gate. Auto-fallback to Standard.');
      const fallback = await generatePDFV2(layout, fieldValues, detailRows, { profile: 'standard' });
      const fbMeta = fallback.meta || {};
      _recordPdfTelemetry({
        type: 'pdf_fallback',
        engine: 'v2',
        fromProfile: requestedProfile,
        toProfile: 'standard',
        layoutId: layout?.id || '',
        layoutName: layout?.name || '',
        reason: 'size_gate_exceeded',
        bytes: fallback.blob?.size || 0,
        pages: fbMeta.pages || 1,
        success: true,
      });
      return fallback.blob;
    }
    if (meta.durationMs > (profileCfg.maxGenerationMs || Number.MAX_SAFE_INTEGER) && requestedProfile === 'print_hd') {
      if (window.showToast) window.showToast('Print HD generation was slow. Auto-fallback to Standard.');
      const fallback = await generatePDFV2(layout, fieldValues, detailRows, { profile: 'standard' });
      _recordPdfTelemetry({
        type: 'pdf_fallback',
        engine: 'v2',
        fromProfile: requestedProfile,
        toProfile: 'standard',
        layoutId: layout?.id || '',
        layoutName: layout?.name || '',
        reason: 'generation_timeout',
        bytes: fallback.blob?.size || 0,
        pages: fallback.meta?.pages || 1,
        success: true,
      });
      return fallback.blob;
    }
    return out.blob;
  } catch (err) {
    _recordPdfTelemetry({
      type: 'pdf_generate',
      engine: 'v2',
      profile: requestedProfile,
      layoutId: layout?.id || '',
      layoutName: layout?.name || '',
      pages: 0,
      detailRows: Array.isArray(detailRows) ? detailRows.length : 0,
      bytes: 0,
      bytesPerPage: 0,
      durationMs: Math.round(performance.now() - startedAt),
      success: false,
      reason: err?.message || 'unknown',
    });
    throw err;
  }
}

/**
 * Render a live preview of the layout into containerEl.
 * Builds the same page DOM as generatePDF but displays it directly (no html2canvas).
 */
async function renderLayoutPreview(layout, fieldValues, detailRows, containerEl) {
  const page = layout.page;
  const plan = _buildCanonicalRenderPlan(layout, fieldValues, detailRows, PDF_MM_TO_PX);
  const { wMm, hMm, pages, totalPages } = plan;

  // Scale to fit container width with a little padding
  const containerWidth = Math.max(200, containerEl.clientWidth - 40);
  const scale = Math.min(PDF_MM_TO_PX, containerWidth / wMm);
  containerEl.innerHTML = '';

  for (let pi = 0; pi < totalPages; pi++) {
    const pagePlan = pages[pi];

    if (totalPages > 1) {
      const lbl = document.createElement('div');
      lbl.className = 'preview-page-label';
      lbl.textContent = `Page ${pi + 1} of ${totalPages}`;
      containerEl.appendChild(lbl);
    }

    const pageDOM = _buildPageDOM(
      page,
      wMm,
      hMm,
      pagePlan.pageEls,
      pagePlan.fieldValues || fieldValues,
      scale,
      pagePlan.detailOverride,
      pagePlan.pageNumber,
      pagePlan.totalPages
    );

    // Zone guide overlays — match designer appearance so positions look identical
      if (_isHeaderActiveOnPage(page, pi)) {
      const hOv = document.createElement('div');
      hOv.style.cssText = `position:absolute;left:0;width:100%;top:0;height:${(page.headerHeight || 20) * scale}px;background:rgba(59,130,246,0.07);border-bottom:2px dashed rgba(59,130,246,0.5);box-sizing:border-box;pointer-events:none;z-index:10;`;
      const hLbl = document.createElement('span');
      hLbl.style.cssText = 'position:absolute;right:8px;bottom:2px;font-size:10px;font-weight:700;letter-spacing:0.08em;color:rgba(59,130,246,0.75);';
      hLbl.textContent = 'HEADER' + (page.headerPages === 'first' ? ' (1st page)' : '');
      hOv.appendChild(hLbl);
      pageDOM.appendChild(hOv);
    }
      if (_isFooterActiveOnPage(page, pi, totalPages)) {
      const footerHeight = (page.footerHeight || 15);
      const fTop = (hMm - footerHeight) * scale;
      const fOv = document.createElement('div');
      fOv.style.cssText = `position:absolute;left:0;width:100%;top:${fTop}px;height:${footerHeight * scale}px;background:rgba(249,115,22,0.07);border-top:2px dashed rgba(249,115,22,0.5);box-sizing:border-box;pointer-events:none;z-index:10;`;
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
window.getLastPdfTelemetry = _getLastPdfTelemetry;
window.getPdfTelemetry = _getPdfTelemetry;
window.clearPdfTelemetry = _clearPdfTelemetry;
window.recordPdfEvent = recordPdfEvent;
window.evaluatePdfReleaseGate = evaluatePdfReleaseGate;
window.checkDeterministicPageBreaks = checkDeterministicPageBreaks;
window.getPdfEmailGuardConfig = function getPdfEmailGuardConfig() {
  const cfg = _getPdfConfig();
  return { ...(cfg.email || {}) };
};


