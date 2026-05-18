import express from 'express'
import authRoutes from './authroutes.js'
import auditRoutes from './Auditroutes.js' 

const router = express.Router()

router.use('/auth', authRoutes)
router.use('/audit', auditRoutes) 
export default router