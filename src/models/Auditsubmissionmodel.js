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
  },
  { _id: false }
);

/* ─────────────────────────────────────────────
   Sub-schema: one section
───────────────────────────────────────────── */
const SectionResultSchema = new mongoose.Schema(
  {
    sectionId:    { type: Number, required: true },
    sectionTitle: { type: String, required: true },
    sectionIcon:  { type: String },
    rows:         [ChecklistRowSchema],
    sectionScore: { type: Number },
    earnedPoints: { type: Number },
    maxPoints:    { type: Number },
    passCount:    { type: Number },
    partialCount: { type: Number },
    failCount:    { type: Number },
    naCount:      { type: Number },
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

    /* ── All 7 sections with every row ── */
    sections: [SectionResultSchema],

    /* ── Customer feedback info (name / mobile / company) ── */
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

    /* ── Rating band ── */
    ratingLabel:  { type: String },
    ratingAction: { type: String },
  },
  {
    timestamps: true,
    collection: 'audit_submissions',
  }
);

export default mongoose.model('AuditSubmission', AuditSubmissionSchema);