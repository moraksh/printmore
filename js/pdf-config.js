'use strict';

(function initPdfConfig() {
  const MB = 1024 * 1024;

  const config = {
    profiles: {
      draft: {
        id: 'draft',
        label: 'Draft (Small)',
        imageJpegQuality: 0.68,
        imageScale: 1.25,
        tableRasterScale: 1.3,
        tableRasterFormat: 'jpeg',
        tableRasterQuality: 0.62,
        barcodeRasterScale: 1.5,
        barcodeScale: 2,
        barcodeModuleWidth: 1.3,
        compress: true,
        maxBytesPerPage: 0.9 * MB,
        maxTotalBytes: 4 * MB,
        maxGenerationMs: 7000,
      },
      standard: {
        id: 'standard',
        label: 'Standard (Balanced)',
        imageJpegQuality: 0.8,
        imageScale: 1.6,
        tableRasterScale: 1.9,
        tableRasterFormat: 'png',
        tableRasterQuality: 0.82,
        barcodeRasterScale: 2,
        barcodeScale: 3,
        barcodeModuleWidth: 1.6,
        compress: true,
        maxBytesPerPage: 1.5 * MB,
        maxTotalBytes: 8 * MB,
        maxGenerationMs: 12000,
      },
      print_hd: {
        id: 'print_hd',
        label: 'Print HD (Sharp)',
        imageJpegQuality: 0.9,
        imageScale: 2,
        tableRasterScale: 2.6,
        tableRasterFormat: 'png',
        tableRasterQuality: 0.92,
        barcodeRasterScale: 2.6,
        barcodeScale: 4,
        barcodeModuleWidth: 2,
        compress: true,
        maxBytesPerPage: 2.3 * MB,
        maxTotalBytes: 14 * MB,
        maxGenerationMs: 18000,
      },
    },
    defaults: {
      engine: 'v2',
      profile: 'standard',
    },
    email: {
      maxAttachmentBytes: 22 * MB,
      hardProviderLimitBytes: 24 * MB,
      timeoutMs: 30000,
    },
    telemetry: {
      maxInMemoryRecords: 100,
      verboseConsole: false,
    },
    releaseGate: {
      maxGenerationMs: {
        draft: 7000,
        standard: 12000,
        print_hd: 18000,
      },
      maxBytesPerPage: {
        draft: 0.9 * MB,
        standard: 1.5 * MB,
        print_hd: 2.3 * MB,
      },
      maxOverlapRegressions: 0,
      barcodeScanSuccessMin: 0.98,
      visualDiffTolerancePct: 1.5,
    },
  };

  window.PRINTMORE_PDF_CONFIG = config;
})();
