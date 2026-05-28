import mongoose from 'mongoose';

/* ─────────────────────────────────────────────
   Sub-schema: one checklist row
   status enum covers both SLA statuses (Pass/Partial/Fail/N.A.)
   AND Customer Feedback ratings (Excellent/Good/Average/Poor)
───────────────────────────────────────────── */
const ChecklistRowSchema = new mongoose.Schema(
  {
    itemIndex: { type: Number, required: true },
    itemText:  { type: String, required: true },
    frequency: { type: String, default: '' },          // ← NOT required (Section 7 has no frequency)
    status: {
      type: String,
      enum: ['Pass', 'Partial', 'Fail', 'N.A.', 'Excellent', 'Good', 'Average', 'Poor'],
      required: true,
    },
    points: { type: Number, required: true },          // 100 / 75 / 50 / 25 / 0
    remark: { type: String, default: '' },
    maxMarks: { type: Number, required: true },        // ← NEW: max marks for this item
  },
  { _id: false }
);

/* ─────────────────────────────────────────────
   Sub-schema: one section
───────────────────────────────────────────── */
const SectionResultSchema = new mongoose.Schema(
  {
    sectionId:         { type: Number, required: true },
    sectionBackendId:  { type: String, default: '' },  // ← NEW: backend category ID for reference
    sectionTitle:      { type: String, required: true },
    sectionIcon:       { type: String },
    sectionType:       { type: String, enum: ['checklist', 'feedback'], default: 'checklist' },  // ← NEW: section type
    rows:              [ChecklistRowSchema],
    sectionScore:      { type: Number },
    earnedPoints:      { type: Number },
    maxPoints:         { type: Number },
    passCount:         { type: Number },
    partialCount:      { type: Number },
    failCount:         { type: Number },
    naCount:           { type: Number },
  },
  { _id: false }
);

/* ─────────────────────────────────────────────
   Sub-schema: customer feedback info (Section 7 header fields)
───────────────────────────────────────────── */
const CustomerFeedbackInfoSchema = new mongoose.Schema(
  {
    name:          { type: String, default: '' },
    mobile:        { type: String, default: '' },
    company:       { type: String, default: '' },
    rating:        { type: Number, min: 0, max: 5, default: 0 },    // ← NEW: Star rating 1-5
    otherComments: { type: String, default: '' },
  },
  { _id: false }
);

/* ─────────────────────────────────────────────
   Main schema
───────────────────────────────────────────── */
const AuditSubmissionSchema = new mongoose.Schema(
  {
    /* ── Who submitted ── */
    auditorName:  { type: String, required: true },
    auditorEmail: { type: String, default: '' },

    /* ── When ── */
    submittedAt: { type: Date, default: Date.now },

    /* ── All sections with every row ── */
    sections: [SectionResultSchema],

    /* ── Customer feedback info (name / mobile / company / rating / otherComments) ── */
    customerFeedback: { type: CustomerFeedbackInfoSchema, default: () => ({}) },

    /* ── Aggregate totals ── */
    overallScore:  { type: Number, required: true },
    earnedMarks:   { type: Number, required: true },
    maxMarks:      { type: Number, required: true },
    passRate:      { type: Number },
    totalItems:    { type: Number },
    totalPass:     { type: Number },
    totalPartial:  { type: Number },
    totalFail:     { type: Number },
    totalNA:       { type: Number },

    /* ── NEW: Additional metrics for admin dashboard ── */
    compliance:      { type: Number, default: 0 },        // ← Compliance percentage (same as overallScore)
    scoredItems:     { type: Number, default: 0 },        // ← Items that have a status (not N.A.)
    criticalFails:   { type: Number, default: 0 },        // ← Count of Fail + Average + Poor items
    
    /* ── Rating band ── */
    ratingLabel:  { type: String },
    ratingAction: { type: String },

    /* ── Period tracking for dashboard filtering ── */
    year:         { type: Number },                        // ← For yearly aggregation
    quarter:      { type: Number },                        // ← Q1, Q2, Q3, Q4
    month:        { type: Number },                        // ← 1-12
    week:         { type: Number },                        // ← ISO week number
    date:         { type: Date, default: Date.now },       // ← Exact date
  },
  {
    timestamps: true,
    collection: 'audit_submissions',
  }
);

// Create indexes for dashboard filtering
AuditSubmissionSchema.index({ overallScore: -1, submittedAt: -1 });  // For top ratings
AuditSubmissionSchema.index({ year: 1, quarter: 1 });               // For quarterly view
AuditSubmissionSchema.index({ year: 1, month: 1 });                 // For monthly view
AuditSubmissionSchema.index({ year: 1, week: 1 });                  // For weekly view
AuditSubmissionSchema.index({ submittedAt: -1 });                   // For recent

export default mongoose.model('AuditSubmission', AuditSubmissionSchema);