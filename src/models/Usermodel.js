import mongoose from 'mongoose'

const userSchema = new mongoose.Schema(
  {
    fname: {
      type: String,
      required: true
    },

    lname: {
      type: String,
      required: true
    },

    email: {
      type: String,
      required: true,
      unique: true
    },

    mobile: {
      type: String,
      required: true
    },

    password: {
      type: String,
      required: true
    },

    role: {
      type: String,
      enum: ['user', 'admin'],
      default: 'user'
    }
  },
  {
    timestamps: true
  }
)

const User = mongoose.model('User', userSchema)

export default User