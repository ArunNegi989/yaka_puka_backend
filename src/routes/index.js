import express from 'express'
import authRoutes from './authroutes.js'
import auditRoutes from './Auditroutes.js'
import slaRoutes from './Slaroutes.js'  

const router = express.Router()

router.use('/auth', authRoutes)
router.use('/audit', auditRoutes)
router.use('/sla', slaRoutes)                     

export default router