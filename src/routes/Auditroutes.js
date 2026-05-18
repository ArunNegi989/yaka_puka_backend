import express from 'express';
import {
  submitAudit,
  getAllAudits,
  getAuditById,
  deleteAudit,
} from '../controllers/Auditcontroller.js';

const router = express.Router();

router.post('/submit',  submitAudit);
router.get('/all',      getAllAudits);
router.get('/:id',      getAuditById);
router.delete('/:id',   deleteAudit);

export default router;