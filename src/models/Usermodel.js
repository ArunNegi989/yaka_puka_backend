import mongoose from 'mongoose'

const userSchema = new mongoose.Schema(
  {
    fname: {
      type: String,
      required: true,
      trim: true
    },

    lname: {
      type: String,
      required: true,
      trim: true
    },

    email: {
      type: String,
      required: true,
      unique: true,
      lowercase: true,
      trim: true
    },

    mobile: {
      type: String,
      required: false,
      trim: true
    },

    password: {
      type: String,
      required: true,
      // ✅ FIX: select:false — password kabhi bhi response mein accidentally nahi aayega
      // Controller mein explicitly .select('+password') likhna padega jab chahiye
      select: false
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