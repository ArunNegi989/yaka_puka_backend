// routes/exportRoutes.js

import express from 'express'
import multer  from 'multer'

import { exportIndividual, exportOverall }                                from '../controllers/Exportcontroller.js'
import { exportDashboardPptx }                                            from '../controllers/PptxExportController.js'

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



export default router