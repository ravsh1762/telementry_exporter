'use strict';
const { Router } = require('express');
const multer = require('multer');
const c = require('../controller/toolsController');

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 150 * 1024 * 1024, files: 500 },
  fileFilter: (_req, file, cb) => {
    if (/\.xlsx$/i.test(file.originalname)) return cb(null, true);
    cb(new Error(`"${file.originalname}" is not an .xlsx file.`));
  },
});

const router = Router();

router.post('/report-merger', upload.array('files'), c.reportMerger);
router.post('/imei-splitter', upload.single('file'), c.imeiSplitter);
router.post('/timestamp-matcher', upload.fields([{ name: 'ex1', maxCount: 1 }, { name: 'ex2', maxCount: 1 }]), c.timestampMatcher);
router.post('/date-filter', upload.single('file'), c.dateFilter);
router.post('/day-shift', upload.single('file'), c.dayShift);
router.post('/last-imei', upload.single('file'), c.lastImei);

module.exports = router;
