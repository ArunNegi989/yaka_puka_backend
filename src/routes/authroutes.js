import express from 'express'

import {
  registerUser,
  loginUser,
  checkAdminExists,
  createUserByAdmin,
  getAllUsers,
  getSingleUser,
  updateUser,
  deleteUser
} from '../controllers/authcontroller.js'

import protect from '../middlewares/authmiddleware.js'
import adminOnly from '../middlewares/rolemiddleware.js'

const router = express.Router()

/* ───────────────── PUBLIC ROUTES ───────────────── */

router.post('/register', registerUser)

router.post('/login', loginUser)

router.get('/check-admin', checkAdminExists)

/* ───────────────── ADMIN CRUD ROUTES ───────────────── */

router.post(
  '/admin/create-user',
  protect,
  adminOnly,
  createUserByAdmin
)

router.get(
  '/admin/users',
  protect,
  adminOnly,
  getAllUsers
)

router.get(
  '/admin/users/:id',
  protect,
  adminOnly,
  getSingleUser
)

router.put(
  '/admin/users/:id',
  protect,
  adminOnly,
  updateUser
)

router.delete(
  '/admin/users/:id',
  protect,
  adminOnly,
  deleteUser
)

export default router