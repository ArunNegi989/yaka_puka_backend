// controllers/Exportcontroller.js

import { generateIndividualExcel, generateOverallExcel } from '../utils/excelGenerators.js'
import Audit from '../models/auditModel.js'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// ── POST /audit/export/individual ────────────────────────────────────────────
export async function exportIndividual(req, res) {
  try {
    const { submission } = req.body
    if (!submission) {
      return res.status(400).json({ error: 'submission is required' })
    }

    const buffer = await generateIndividualExcel(submission)

    const auditorName = (submission.auditorName ?? 'audit').replace(/\s+/g, '_')
    const date        = new Date(submission.submittedAt ?? Date.now()).toISOString().slice(0, 10)
    const filename    = `SLA_Audit_${auditorName}_${date}.xlsx`

    res.setHeader('Content-Type', XLSX_MIME)
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.setHeader('Content-Length', buffer.length)
    res.send(buffer)
  } catch (err) {
    console.error('[exportController] individual export error:', err)
    res.status(500).json({ error: 'Failed to generate Excel report', details: err.message })
  }
}

// ── POST /audit/export/overall ───────────────────────────────────────────────
// OPTION 1: DB se seedha fetch karta hai
export async function exportOverall(req, res) {
  try {
    // Try DB first
    let submissions = await Audit.find().lean()
    
    // If DB empty, try frontend data
    if (!submissions || submissions.length === 0) {
      const { submissions: frontendData } = req.body
      if (frontendData && Array.isArray(frontendData) && frontendData.length > 0) {
        submissions = frontendData
      } else {
        return res.status(404).json({ error: 'No submissions found' })
      }
    }

    const buffer   = await generateOverallExcel(submissions)
    const date     = new Date().toISOString().slice(0, 10)
    const filename = `SLA_Audit_Overall_${date}.xlsx`

    res.setHeader('Content-Type', XLSX_MIME)
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.setHeader('Content-Length', buffer.length)
    res.send(buffer)
  } catch (err) {
    console.error('[exportController] overall export error:', err)
    res.status(500).json({ error: 'Failed to generate Excel report', details: err.message })
  }
}