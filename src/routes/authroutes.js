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

const router = express.Router()

/* AUTH */
router.post('/register', registerUser)
router.post('/login', loginUser)

/* ADMIN */
router.get('/check-admin', checkAdminExists)

router.post('/admin/create-user', createUserByAdmin)

router.get('/admin/users', getAllUsers)

router.get('/admin/user/:id', getSingleUser)

router.put('/admin/user/:id', updateUser)

router.delete('/admin/user/:id', deleteUser)

export default router