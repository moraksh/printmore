'use strict';

const fs = require('fs');
const path = require('path');

function loadJson(p) {
  return JSON.parse(fs.readFileSync(p, 'utf8'));
}

function main() {
  const baselinePath = process.argv[2] || path.join(__dirname, '..', 'tests', 'pdf-v2-golden-baseline.json');
  const currentPath = process.argv[3] || path.join(__dirname, '..', 'tests', 'pdf-v2-current-metrics.sample.json');

  const baseline = loadJson(baselinePath);
  const current = loadJson(currentPath);
  const currentMap = new Map((current.runs || []).map(r => [String(r.id), r]));

  const issues = [];
  (baseline.cases || []).forEach(tc => {
    const run = currentMap.get(String(tc.id));
    if (!run) {
      issues.push(`Missing run for ${tc.id}`);
      return;
    }
    if ((run.bytes || 0) > (tc.maxBytes || Number.MAX_SAFE_INTEGER)) {
      issues.push(`${tc.id}: bytes ${run.bytes} > ${tc.maxBytes}`);
    }
    if ((run.durationMs || 0) > (tc.maxMs || Number.MAX_SAFE_INTEGER)) {
      issues.push(`${tc.id}: duration ${run.durationMs}ms > ${tc.maxMs}ms`);
    }
    if (run.overlapRegression) {
      issues.push(`${tc.id}: overlap regression detected`);
    }
    if (run.barcodeScanSuccess === false) {
      issues.push(`${tc.id}: barcode scan failed`);
    }
    if (Number(run.visualDiffPct || 0) > 1.5) {
      issues.push(`${tc.id}: visual diff ${run.visualDiffPct}% > 1.5%`);
    }
  });

  if (issues.length) {
    console.error('PDF V2 release gate: FAILED');
    issues.forEach(i => console.error(` - ${i}`));
    process.exit(1);
  }
  console.log('PDF V2 release gate: PASSED');
}

main();

