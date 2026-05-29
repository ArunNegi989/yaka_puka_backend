// routes/index.js

import express from 'express'
import authRoutes   from './authroutes.js'
import auditRoutes  from './Auditroutes.js'
import slaRoutes    from './Slaroutes.js'
import exportRoutes from './exportRoutes.js'

const router = express.Router()

router.use('/auth',  authRoutes)
router.use('/sla',   slaRoutes)

const auditRouter = express.Router()
auditRouter.use('/', auditRoutes)
auditRouter.use('/', exportRoutes)
router.use('/audit', auditRouter)

export default router