import User from '../models/Usermodel.js'
import bcrypt from 'bcryptjs'
import generateToken from '../utils/generateToken.js'

/* ───────────────── REGISTER USER ───────────────── */

export const registerUser = async (req, res) => {
  try {
    const {
      fname,
      lname,
      email,
      mobile,
      password
    } = req.body

    const existingUser = await User.findOne({ email })

    if (existingUser) {
      return res.status(400).json({
        success: false,
        message: 'User already exists'
      })
    }

    const hashedPassword = await bcrypt.hash(password, 10)

    const user = await User.create({
      fname,
      lname,
      email,
      mobile,
      password: hashedPassword
    })

    res.status(201).json({
      success: true,
      message: 'User Registered Successfully',
      token: generateToken(user._id, user.role),
      user
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}

/* ───────────────── LOGIN USER ───────────────── */

export const loginUser = async (req, res) => {
  try {
    const { email, password } = req.body

    const user = await User.findOne({ email })

    if (!user) {
      return res.status(400).json({
        success: false,
        message: 'Invalid Credentials'
      })
    }

    const isMatch = await bcrypt.compare(
      password,
      user.password
    )

    if (!isMatch) {
      return res.status(400).json({
        success: false,
        message: 'Invalid Credentials'
      })
    }

    res.status(200).json({
      success: true,
      message: 'Login Successful',
      token: generateToken(user._id, user.role),
      user
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}

/* ───────────────── CHECK ADMIN EXISTS ───────────────── */

export const checkAdminExists = async (req, res) => {
  try {
    const admin = await User.findOne({
      role: 'admin'
    })

    res.status(200).json({
      success: true,
      adminExists: !!admin
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}

/* ───────────────── CREATE USER BY ADMIN ───────────────── */

export const createUserByAdmin = async (req, res) => {
  try {
    const {
      fname,
      lname,
      email,
      mobile,
      password
    } = req.body

    const existingUser = await User.findOne({ email })

    if (existingUser) {
      return res.status(400).json({
        success: false,
        message: 'User already exists'
      })
    }

    const hashedPassword = await bcrypt.hash(password, 10)

    const user = await User.create({
      fname,
      lname,
      email,
      mobile,
      password: hashedPassword,
      role: 'user'
    })

    res.status(201).json({
      success: true,
      message: 'User created successfully',
      user
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}

/* ───────────────── GET ALL USERS ───────────────── */

export const getAllUsers = async (req, res) => {
  try {
    const users = await User.find({
      role: 'user'
    })
      .select('-password')
      .sort({ createdAt: -1 })

    res.status(200).json({
      success: true,
      users
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}

/* ───────────────── GET SINGLE USER ───────────────── */

export const getSingleUser = async (req, res) => {
  try {
    const user = await User.findById(req.params.id)
      .select('-password')

    if (!user) {
      return res.status(404).json({
        success: false,
        message: 'User not found'
      })
    }

    res.status(200).json({
      success: true,
      user
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}


/* ───────────────── UPDATE USER ───────────────── */

export const updateUser = async (req, res) => {
  try {
    const {
      fname,
      lname,
      mobile,
      password        // optional — only if admin wants to change password
    } = req.body

    const user = await User.findById(req.params.id)

    if (!user) {
      return res.status(404).json({
        success: false,
        message: 'User not found'
      })
    }

    // Update allowed fields (email is intentionally excluded)
    user.fname  = fname  || user.fname
    user.lname  = lname  || user.lname
    user.mobile = mobile || user.mobile

    // Update password only if provided
    if (password) {
      const bcrypt = await import('bcryptjs')
      user.password = await bcrypt.hash(password, 10)
    }

    const updatedUser = await user.save()

    res.status(200).json({
      success: true,
      message: 'User updated successfully',
      user: updatedUser
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}
/* ───────────────── DELETE USER ───────────────── */

export const deleteUser = async (req, res) => {
  try {
    const user = await User.findById(req.params.id)

    if (!user) {
      return res.status(404).json({
        success: false,
        message: 'User not found'
      })
    }

    await User.findByIdAndDelete(req.params.id)

    res.status(200).json({
      success: true,
      message: 'User deleted successfully'
    })

  } catch (error) {
    res.status(500).json({
      success: false,
      message: error.message
    })
  }
}