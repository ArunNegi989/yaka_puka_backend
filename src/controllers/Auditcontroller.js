import AuditSubmission from '../models/Auditsubmissionmodel.js';

/* ══════════════════════════════════════════════
   POST /api/audit/submit
══════════════════════════════════════════════ */
export const submitAudit = async (req, res) => {
  try {
    const {
      auditorName,
      auditorEmail,
      sections,
      customerFeedback,       // ← name / mobile / company / otherComments
      overallScore,
      earnedMarks,
      maxMarks,
      passRate,
      totalItems,
      totalPass,
      totalPartial,
      totalFail,
      totalNA,
      ratingLabel,
      ratingAction,
    } = req.body;

    if (!auditorName || !sections || overallScore === undefined) {
      return res.status(400).json({ message: 'Missing required fields.' });
    }

    const submission = new AuditSubmission({
      auditorName,
      auditorEmail:     auditorEmail     || '',
      sections,
      customerFeedback: customerFeedback || {},
      overallScore,
      earnedMarks,
      maxMarks,
      passRate,
      totalItems,
      totalPass,
      totalPartial,
      totalFail,
      totalNA,
      ratingLabel,
      ratingAction,
    });

    await submission.save();

    return res.status(201).json({
      message: 'Audit submitted successfully.',
      id: submission._id,
    });
  } catch (err) {
    console.error('Audit submit error:', err);
    return res.status(500).json({ message: 'Server error. Please try again.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/all
══════════════════════════════════════════════ */
export const getAllAudits = async (req, res) => {
  try {
    const page  = parseInt(req.query.page)  || 1;
    const limit = parseInt(req.query.limit) || 50;
    const skip  = (page - 1) * limit;

    const [submissions, total] = await Promise.all([
      AuditSubmission.find({})
        .sort({ submittedAt: -1 })
        .skip(skip)
        .limit(limit)
        .select('-__v'),
      AuditSubmission.countDocuments(),
    ]);

    return res.json({
      total,
      page,
      pages: Math.ceil(total / limit),
      submissions,
    });
  } catch (err) {
    console.error('Audit list error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/:id
══════════════════════════════════════════════ */
export const getAuditById = async (req, res) => {
  try {
    const submission = await AuditSubmission.findById(req.params.id).select('-__v');
    if (!submission) {
      return res.status(404).json({ message: 'Submission not found.' });
    }
    return res.json(submission);
  } catch (err) {
    console.error('Audit fetch error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   DELETE /api/audit/:id
══════════════════════════════════════════════ */
export const deleteAudit = async (req, res) => {
  try {
    const deleted = await AuditSubmission.findByIdAndDelete(req.params.id);
    if (!deleted) {
      return res.status(404).json({ message: 'Submission not found.' });
    }
    return res.json({ message: 'Submission deleted successfully.' });
  } catch (err) {
    console.error('Audit delete error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};