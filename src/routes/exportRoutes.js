// routes/exportRoutes.js

import express from 'express'
import multer  from 'multer'

import { exportIndividual, exportOverall }                                from '../controllers/Exportcontroller.js'
import { exportDashboardPptx }                                            from '../controllers/PptxExportController.js'
import { exportDashboardPptxScreenshot, screenshotUpload }                from '../controllers/PptxScreenshotController.js'

const router = express.Router()

// ── existing excel exports ─────────────────────────────────────────────────────
// POST /audit/export/individual
router.post('/export/individual', exportIndividual)

// POST /audit/export/overall
router.post('/export/overall', exportOverall)

// ── new dashboard exports ──────────────────────────────────────────────────────
// POST /audit/export/dashboard-pptx
// Generates a full-featured PPTX with charts, graphs, tables from DB
router.post('/export/dashboard-pptx', exportDashboardPptx)

// POST /audit/export/dashboard-pptx-screenshot
// Accepts a PNG screenshot from frontend and wraps into PPTX
router.post(
  '/export/dashboard-pptx-screenshot',
  screenshotUpload.single('screenshot'),
  exportDashboardPptxScreenshot,
)

export default router