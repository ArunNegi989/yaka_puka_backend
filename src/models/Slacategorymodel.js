import mongoose from "mongoose";

/* ─── Question Sub-Schema ─── */
const SlaQuestionSchema = new mongoose.Schema(
  {
    id: {
      type: String,
      required: true,
    },
    text: {
      type: String,
      required: [true, "Question text is required"],
      trim: true,
    },
    frequency: {
      type: String,
      required: true,
      default: "Daily",
      trim: true,
    },
    maxMarks: {
      type: Number,
      required: true,
      default: 100,
      min: [1, "Max marks must be at least 1"],
      max: [1000, "Max marks cannot exceed 1000"],
    },
    required: {
      type: Boolean,
      default: true,
    },
    order: {
      type: Number,
      default: 0,
    },
  },
  { _id: false } // use the frontend-generated id field instead
);

/* ─── Category Schema ─── */
const SlaCategorySchema = new mongoose.Schema(
  {
    id: {
      type: String,
      required: true,
      unique: true,
    },
    title: {
      type: String,
      required: [true, "Category title is required"],
      trim: true,
    },
    icon: {
      type: String,
      default: "📋",
    },
    type: {
      type: String,
      enum: {
        values: ["checklist", "feedback"],
        message: "Type must be either 'checklist' or 'feedback'",
      },
      default: "checklist",
    },
    order: {
      type: Number,
      default: 0,
    },
    isActive: {
      type: Boolean,
      default: true,
    },
    questions: {
      type: [SlaQuestionSchema],
      default: [],
    },
  },
  {
    timestamps: true, // createdAt, updatedAt
    versionKey: false,
  }
);

/* ─── Indexes ─── */
SlaCategorySchema.index({ order: 1 });
SlaCategorySchema.index({ isActive: 1 });

const SlaCategory = mongoose.model("SlaCategory", SlaCategorySchema);

export default SlaCategory;