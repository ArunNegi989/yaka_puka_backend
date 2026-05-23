// routes/exportRoutes.js

import express from 'express'
import { exportIndividual, exportOverall } from '../controllers/Exportcontroller.js'

const router = express.Router()

// POST /audit/export/individual
router.post('/export/individual', exportIndividual)

// POST /audit/export/overall
router.post('/export/overall', exportOverall)

export default router