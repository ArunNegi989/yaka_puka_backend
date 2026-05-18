import mongoose from 'mongoose';

/* ─────────────────────────────────────────────
   Sub-schema: one checklist row
───────────────────────────────────────────── */
const ChecklistRowSchema = new mongoose.Schema(
  {
    itemIndex:   { type: Number, required: true },        // 0-based index within section
    itemText:    { type: String, required: true },        // full checklist text
    frequency:   { type: String, required: true },        // e.g. "Daily", "Monthly"
    status:      { type: String, enum: ['Pass', 'Partial', 'Fail', 'N.A.'], required: true },
    points:      { type: Number, required: true },        // 100 / 50 / 0
    remark:      { type: String, default: '' },
  },
  { _id: false }
);

/* ─────────────────────────────────────────────
   Sub-schema: one section
───────────────────────────────────────────── */
const SectionResultSchema = new mongoose.Schema(
  {
    sectionId:    { type: Number, required: true },       // 1–6
    sectionTitle: { type: String, required: true },
    sectionIcon:  { type: String },
    rows:         [ChecklistRowSchema],
    sectionScore: { type: Number },                       // % score for this section
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
   Main schema
───────────────────────────────────────────── */
const AuditSubmissionSchema = new mongoose.Schema(
  {
    /* ── Who submitted ── */
    auditorName:  { type: String, required: true },
    auditorEmail: { type: String, default: '' },

    /* ── When ── */
    submittedAt:  { type: Date, default: Date.now },

    /* ── All 6 sections with every row ── */
    sections: [SectionResultSchema],

    /* ── Aggregate totals ── */
    overallScore:    { type: Number, required: true },   // overall % (0-100)
    earnedMarks:     { type: Number, required: true },   // e.g. 4750
    maxMarks:        { type: Number, required: true },   // e.g. 5400
    passRate:        { type: Number },                   // % of items that passed
    totalItems:      { type: Number },
    totalPass:       { type: Number },
    totalPartial:    { type: Number },
    totalFail:       { type: Number },
    totalNA:         { type: Number },

    /* ── Rating band ── */
    ratingLabel:  { type: String },                      // "Excellent" / "Good" / etc.
    ratingAction: { type: String },
  },
  {
    timestamps: true,   // adds createdAt & updatedAt
    collection: 'audit_submissions',
  }
);

export default mongoose.model('AuditSubmission', AuditSubmissionSchema);