import express from 'express';
import {
  submitAudit,
  getAllAudits,
  getAuditById,
  deleteAudit,
  getRecentAudits,
  getTopRatingsAudits,
  getTopCustomerRatingsAudits,
  getDashboardStats,
  getDashboardTrend,
  getCriticalItems,
  getTopItemsByRating,
  getSectionItems,
} from '../controllers/Auditcontroller.js';

const router = express.Router();

/* ── Core CRUD ── */
router.post('/submit',   submitAudit);
router.get('/all',       getAllAudits);
router.get('/recent',    getRecentAudits);

/* ── Dashboard endpoints ── */
// IMPORTANT: all /dashboard/* routes must come BEFORE /:id
router.get('/dashboard/top-ratings',          getTopRatingsAudits);
router.get('/dashboard/top-customer-ratings', getTopCustomerRatingsAudits);
router.get('/dashboard/stats',                getDashboardStats);
router.get('/dashboard/trend',                getDashboardTrend);
router.get('/dashboard/critical-items',       getCriticalItems);
router.get('/dashboard/top-items',            getTopItemsByRating);
router.get('/dashboard/section-items',        getSectionItems);

/* ── Single record routes (must be last) ── */
router.get('/:id',    getAuditById);
router.delete('/:id', deleteAudit);

export default router;