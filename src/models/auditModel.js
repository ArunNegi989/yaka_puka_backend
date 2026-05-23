// models/auditModel.js

import mongoose from 'mongoose'

const ChecklistRowSchema = new mongoose.Schema({
  itemIndex: { type: Number },
  itemText:  { type: String },
  frequency: { type: String },
  status:    { type: String, enum: ['Pass', 'Partial', 'Fail', 'N.A.'] },
  points:    { type: Number },
  remark:    { type: String, default: '' },
}, { _id: false })

const SectionResultSchema = new mongoose.Schema({
  sectionId:    { type: Number },
  sectionTitle: { type: String },
  sectionIcon:  { type: String },
  rows:         [ChecklistRowSchema],
  sectionScore: { type: Number },
  earnedPoints: { type: Number },
  maxPoints:    { type: Number },
  passCount:    { type: Number },
  partialCount: { type: Number },
  failCount:    { type: Number },
  naCount:      { type: Number },
}, { _id: false })

const CustomerFeedbackSchema = new mongoose.Schema({
  name:          { type: String, default: '' },
  mobile:        { type: String, default: '' },
  company:       { type: String, default: '' },
  otherComments: { type: String, default: '' },
  auditorEmail:  { type: String, default: '' },
}, { _id: false })

const AuditSchema = new mongoose.Schema({
  auditorName:  { type: String, required: true },
  auditorEmail: { type: String, default: '' },
  submittedAt:  { type: Date,   default: Date.now },

  sections:         [SectionResultSchema],
  customerFeedback: CustomerFeedbackSchema,

  overallScore: { type: Number, default: 0 },
  earnedMarks:  { type: Number, default: 0 },
  maxMarks:     { type: Number, default: 0 },
  passRate:     { type: Number, default: 0 },

  totalItems:   { type: Number, default: 0 },
  totalPass:    { type: Number, default: 0 },
  totalPartial: { type: Number, default: 0 },
  totalFail:    { type: Number, default: 0 },
  totalNA:      { type: Number, default: 0 },

  ratingLabel:  { type: String, default: '' },
  ratingAction: { type: String, default: '' },
}, {
  timestamps: true,
})

const Audit = mongoose.model('Audit', AuditSchema)

export default Audit