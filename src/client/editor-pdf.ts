import * as pdfjs from 'pdfjs-dist';

pdfjs.GlobalWorkerOptions.workerSrc = '/js/pdf.worker.mjs';
console.info('pdf.js', pdfjs.version);
