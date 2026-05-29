import AuditSubmission from '../models/Auditsubmissionmodel.js';

/* ─────────────────────────────────────────────────────────
   HELPER: Calculate period info (year, quarter, month, week)
─────────────────────────────────────────────────────────── */
function getPeriodInfo(date = new Date()) {
  const d       = new Date(date);
  const year    = d.getFullYear();
  const month   = d.getMonth() + 1;
  const quarter = Math.ceil(month / 3);

  const firstDay = new Date(year, 0, 1);
  const daysDiff = Math.floor((d - firstDay) / (24 * 60 * 60 * 1000));
  const week     = Math.ceil((daysDiff + firstDay.getDay() + 1) / 7);

  return { year, month, quarter, week };
}

/* ─────────────────────────────────────────────────────────
   HELPER: Build period match filter from query params
─────────────────────────────────────────────────────────── */
function buildPeriodFilter(query) {
  const period  = query.period || 'all';
  const year    = parseInt(query.year)    || new Date().getFullYear();
  const quarter = parseInt(query.quarter);
  const month   = parseInt(query.month);
  const week    = parseInt(query.week);

  let filter = {};
  if      (period === 'annual'    || period === 'yearly')   filter = { year };
  else if (period === 'quarterly')                           filter = { year, quarter };
  else if (period === 'monthly')                             filter = { year, month };
  else if (period === 'weekly')                              filter = { year, week };

  return { period, filter };
}

/* ─────────────────────────────────────────────────────────
   SCORING HELPER
   • Checklist: Pass = full maxMarks, Partial = maxMarks/2,
                Fail = 0, N.A. = 0 (and excluded from max pool)
   • Feedback:  no marks — review only
                points = 0, maxMarks contribution = 0
─────────────────────────────────────────────────────────── */
function computeSectionScores(sections) {
  let grandPts = 0, grandMax = 0;
  let totalPass = 0, totalPartial = 0, totalFail = 0, totalNA = 0;

  const processedSections = sections.map((sec) => {
    const isFeedback = sec.sectionType === 'feedback';

    let earnedPoints = 0;
    let effectiveMax = 0;
    let passCount    = 0;
    let partialCount = 0;
    let failCount    = 0;
    let naCount      = 0;

    const rows = (sec.rows || []).map((row) => {
      const status  = row.status || (isFeedback ? 'Average' : 'N.A.');
      const itemMax = row.maxMarks ?? 100;

      let points = 0;

      if (!isFeedback) {
        /* ── Checklist scoring ── */
        if (status === 'Pass') {
          points       = itemMax;
          earnedPoints += itemMax;
          effectiveMax += itemMax;   // N.A. items NOT added to max
          passCount++;
        } else if (status === 'Partial') {
          points        = itemMax / 2;
          earnedPoints += itemMax / 2;
          effectiveMax += itemMax;
          partialCount++;
        } else if (status === 'Fail') {
          points        = 0;
          earnedPoints += 0;
          effectiveMax += itemMax;   // Fail IS in max pool
          failCount++;
        } else if (status === 'N.A.') {
          points  = 0;
          naCount++;
          // N.A.: NOT added to earnedPoints or effectiveMax
        }
      } else {
        /* ── Feedback: review only, no marks ── */
        points = 0;
        if      (status === 'Excellent') passCount++;
        else if (status === 'Good')      partialCount++;
        else if (status === 'Average')   failCount++;
        else if (status === 'Poor')      naCount++;
      }

      return { ...row, status, points, maxMarks: isFeedback ? 0 : itemMax };
    });

    const sectionScore = effectiveMax > 0
      ? Math.round((earnedPoints / effectiveMax) * 100)
      : 0;

    if (!isFeedback) {
      grandPts     += earnedPoints;
      grandMax     += effectiveMax;
      totalPass    += passCount;
      totalPartial += partialCount;
      totalFail    += failCount;
      totalNA      += naCount;
    }

    return {
      ...sec,
      rows,
      sectionScore,
      earnedPoints,
      maxPoints: effectiveMax,
      passCount,
      partialCount,
      failCount,
      naCount,
    };
  });

  const overallScore = grandMax > 0 ? Math.round((grandPts / grandMax) * 100) : 0;

  return {
    processedSections,
    grandPts,
    grandMax,
    overallScore,
    totalPass,
    totalPartial,
    totalFail,
    totalNA,
  };
}

/* ══════════════════════════════════════════════
   POST /api/audit/submit
══════════════════════════════════════════════ */
export const submitAudit = async (req, res) => {
  try {
    const {
      auditorName,
      auditorEmail,
      sections,
      customerFeedback,
      /* The frontend sends pre-computed values; we re-compute server-side
         to guarantee correctness and prevent tampering. */
      totalItems,
      ratingLabel,
      ratingAction,
    } = req.body;

    if (!auditorName || !sections || !Array.isArray(sections)) {
      return res.status(400).json({ message: 'Missing required fields.' });
    }

    /* ── Re-compute all scores server-side ── */
    const {
      processedSections,
      grandPts,
      grandMax,
      overallScore,
      totalPass,
      totalPartial,
      totalFail,
      totalNA,
    } = computeSectionScores(sections);

    const checklistSections = processedSections.filter((s) => s.sectionType !== 'feedback');
    const applicableItems =
  totalPass + totalPartial + totalFail;

const passRate =
  applicableItems > 0
    ? Math.round((totalPass / applicableItems) * 100)
    : 0;

    const scoredItems = applicableItems;
    const criticalFails = totalFail;    // Fail count from checklist only
    const compliance    = overallScore; // Same as overall

    const SCORE_RANGES = [
      { min: 95, max: 100, label: 'Excellent',         action: 'Maintain current standards'    },
      { min: 85, max: 94,  label: 'Good',              action: 'Monitor and sustain'            },
      { min: 70, max: 84,  label: 'Needs Improvement', action: 'Root cause analysis required'   },
      { min: 0,  max: 69,  label: 'Critical',          action: 'Immediate remediation required' },
    ];
    const range = SCORE_RANGES.find((r) => overallScore >= r.min) ?? SCORE_RANGES[3];

    const periodInfo = getPeriodInfo();
const totalChecklistItems = checklistSections.reduce(
  (a, s) => a + (s.rows || []).length,
  0
);
    const submission = new AuditSubmission({
      auditorName,
      auditorEmail:     auditorEmail     || '',
      sections:         processedSections,
      customerFeedback: customerFeedback || {},

      overallScore,
      earnedMarks:   Math.round(grandPts * 10) / 10,
      maxMarks:      grandMax,
      passRate,
      totalItems: totalChecklistItems,
      totalPass,
      totalPartial,
      totalFail,
      totalNA,

      compliance,
      scoredItems,
      criticalFails,

      ratingLabel:  ratingLabel  || range.label,
      ratingAction: ratingAction || range.action,

      ...periodInfo,
      date: new Date(),
    });

    await submission.save();

    return res.status(201).json({
      message:     'Audit submitted successfully.',
      id:          submission._id,
      overallScore,
      earnedMarks: Math.round(grandPts * 10) / 10,
      maxMarks:    grandMax,
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
   GET /api/audit/recent — last 8 submissions
══════════════════════════════════════════════ */
export const getRecentAudits = async (req, res) => {
  try {
    const submissions = await AuditSubmission.find({})
      .sort({ submittedAt: -1 })
      .limit(8)
      .select('auditorName auditorEmail overallScore ratingLabel submittedAt customerFeedback compliance');
    return res.json({ submissions });
  } catch (err) {
    console.error('Recent audits error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/dashboard/top-ratings
══════════════════════════════════════════════ */
export const getTopRatingsAudits = async (req, res) => {
  try {
    const { period, filter } = buildPeriodFilter(req.query);

    const topAudits = await AuditSubmission.find(filter)
      .sort({ overallScore: -1, submittedAt: -1 })
      .limit(5)
      .select('auditorName overallScore ratingLabel customerFeedback.rating customerFeedback.name submittedAt');

    return res.json({ period, filter, topAudits, count: topAudits.length });
  } catch (err) {
    console.error('Top ratings error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/dashboard/top-customer-ratings
══════════════════════════════════════════════ */
export const getTopCustomerRatingsAudits = async (req, res) => {
  try {
    const { period, filter } = buildPeriodFilter(req.query);

    const topAudits = await AuditSubmission.find(filter)
      .sort({ 'customerFeedback.rating': -1, overallScore: -1, submittedAt: -1 })
      .limit(5)
      .select('auditorName overallScore customerFeedback submittedAt');

    return res.json({ period, filter, topAudits, count: topAudits.length });
  } catch (err) {
    console.error('Top customer ratings error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/dashboard/stats
══════════════════════════════════════════════ */
export const getDashboardStats = async (req, res) => {
  try {
    const { period, filter } = buildPeriodFilter(req.query);

    const stats = await AuditSubmission.aggregate([
      { $match: filter },
      {
        $group: {
          _id: null,
          totalAudits:        { $sum: 1 },
          avgScore:           { $avg: '$overallScore' },
          avgCompliance:      { $avg: '$compliance' },
          avgCustomerRating:  { $avg: '$customerFeedback.rating' },
          totalPass:          { $sum: '$totalPass' },
          totalFail:          { $sum: '$totalFail' },
          totalPartial:       { $sum: '$totalPartial' },
          totalNA:            { $sum: '$totalNA' },
          totalCriticalFails: { $sum: '$criticalFails' },
          highestScore:       { $max: '$overallScore' },
          lowestScore:        { $min: '$overallScore' },
        },
      },
      {
        $project: {
          _id: 0,
          totalAudits:        1,
          avgScore:           { $round: ['$avgScore',          2] },
          avgCompliance:      { $round: ['$avgCompliance',     2] },
          avgCustomerRating:  { $round: ['$avgCustomerRating', 2] },
          totalPass:          1,
          totalFail:          1,
          totalPartial:       1,
          totalNA:            1,
          totalCriticalFails: 1,
          highestScore:       1,
          lowestScore:        1,
        },
      },
    ]);

    return res.json({
      period,
      filter,
      stats: stats[0] || {
        totalAudits: 0, avgScore: 0, avgCompliance: 0, avgCustomerRating: 0,
        totalPass: 0, totalFail: 0, totalPartial: 0, totalNA: 0,
        totalCriticalFails: 0, highestScore: 0, lowestScore: 0,
      },
    });
  } catch (err) {
    console.error('Dashboard stats error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/dashboard/trend
══════════════════════════════════════════════ */
export const getDashboardTrend = async (req, res) => {
  try {
    const groupBy = req.query.groupBy || 'daily';

    let groupStage = {};
    if      (groupBy === 'daily')   groupStage = { $dateToString: { format: '%Y-%m-%d', date: '$submittedAt' } };
    else if (groupBy === 'weekly')  groupStage = { $dateToString: { format: '%Y-W%V',   date: '$submittedAt' } };
    else if (groupBy === 'monthly') groupStage = { $dateToString: { format: '%Y-%m',    date: '$submittedAt' } };

    const trend = await AuditSubmission.aggregate([
      {
        $group: {
          _id:               groupStage,
          count:             { $sum: 1 },
          avgScore:          { $avg: '$overallScore' },
          avgCustomerRating: { $avg: '$customerFeedback.rating' },
          totalFails:        { $sum: '$totalFail' },
        },
      },
      { $sort: { _id: 1 } },
    ]);

    return res.json({ groupBy, trend });
  } catch (err) {
    console.error('Dashboard trend error:', err);
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

/* ══════════════════════════════════════════════
   GET /api/audit/dashboard/critical-items
   Only Fail rows from checklist sections
══════════════════════════════════════════════ */
export const getCriticalItems = async (req, res) => {
  try {
    const { period, filter } = buildPeriodFilter(req.query);
    const limit = parseInt(req.query.limit) || 20;

    const result = await AuditSubmission.aggregate([
      { $match: filter },
      { $unwind: { path: '$sections',      preserveNullAndEmpty: false } },
      /* Exclude feedback sections from critical items */
      { $match: { 'sections.sectionType': { $ne: 'feedback' } } },
      { $unwind: { path: '$sections.rows', preserveNullAndEmpty: false } },
      /* Only Fail rows */
      { $match: { 'sections.rows.status': 'Fail' } },
      {
        $group: {
          _id:        '$sections.rows.itemText',
          count:      { $sum: 1 },
          categories: { $addToSet: '$sections.sectionTitle' },
          auditors:   { $addToSet: '$auditorName' },
          lastSeen:   { $max: '$submittedAt' },
        },
      },
      { $sort: { count: -1 } },
      { $limit: limit },
      {
        $project: {
          _id: 0,
          itemText:     '$_id',
          count:        1,
          categories:   1,
          auditorCount: { $size: '$auditors' },
          lastSeen:     1,
        },
      },
    ]);

    return res.json({ period, total: result.length, criticalItems: result });
  } catch (err) {
    console.error('Critical items error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/dashboard/top-items
   Excludes feedback sections from all buckets
══════════════════════════════════════════════ */
export const getTopItemsByRating = async (req, res) => {
  try {
    const { period, filter } = buildPeriodFilter(req.query);
    const topN = parseInt(req.query.topN) || 5;

    const SCORE_RANGES = [
      { label: 'Excellent',         min: 95, max: 100, color: '#22c55e' },
      { label: 'Good',              min: 85, max: 94,  color: '#3b82f6' },
      { label: 'Needs Improvement', min: 70, max: 84,  color: '#f59e0b' },
      { label: 'Critical',          min: 0,  max: 69,  color: '#ef4444' },
    ];

    const buckets = {};

    for (const range of SCORE_RANGES) {
      const result = await AuditSubmission.aggregate([
        {
          $match: {
            ...filter,
            overallScore: { $gte: range.min, $lte: range.max },
          },
        },
        { $unwind: { path: '$sections',      preserveNullAndEmpty: false } },
        /* Skip feedback sections */
        { $match: { 'sections.sectionType': { $ne: 'feedback' } } },
        { $unwind: { path: '$sections.rows', preserveNullAndEmpty: false } },
        {
          $group: {
            _id:          '$sections.rows.itemText',
            totalCount:   { $sum: 1 },
            passCount:    { $sum: { $cond: [{ $eq: ['$sections.rows.status', 'Pass']    }, 1, 0] } },
            failCount:    { $sum: { $cond: [{ $eq: ['$sections.rows.status', 'Fail']    }, 1, 0] } },
            partialCount: { $sum: { $cond: [{ $eq: ['$sections.rows.status', 'Partial'] }, 1, 0] } },
          },
        },
        {
          $addFields: {
            passRate: {
              $round: [
                { $multiply: [{ $divide: ['$passCount', { $max: ['$totalCount', 1] }] }, 100] },
                1,
              ],
            },
          },
        },
        { $sort: { passCount: -1, failCount: 1 } },
        { $limit: topN },
        {
          $project: {
            _id: 0,
            itemText:     '$_id',
            totalCount:   1,
            passCount:    1,
            failCount:    1,
            partialCount: 1,
            passRate:     1,
          },
        },
      ]);

      buckets[range.label] = { ...range, items: result };
    }

    return res.json({ period, topN, buckets });
  } catch (err) {
    console.error('Top items by rating error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};

/* ══════════════════════════════════════════════
   GET /api/audit/dashboard/section-items
   Per-section breakdown (feedback sections tagged separately)
══════════════════════════════════════════════ */
export const getSectionItems = async (req, res) => {
  try {
    const { period, filter } = buildPeriodFilter(req.query);
    const sectionTitle = req.query.sectionTitle;

    const result = await AuditSubmission.aggregate([
      { $match: filter },
      { $unwind: { path: '$sections', preserveNullAndEmpty: false } },
      ...(sectionTitle ? [{ $match: { 'sections.sectionTitle': sectionTitle } }] : []),
      { $unwind: { path: '$sections.rows', preserveNullAndEmpty: false } },

      /* For checklist: only Pass/Fail/Partial. For feedback: all statuses */
      {
        $match: {
          $or: [
            { 'sections.sectionType': 'feedback' },
            { 'sections.rows.status': { $in: ['Pass', 'Fail', 'Partial'] } },
          ],
        },
      },

      {
        $group: {
          _id: {
            sectionTitle: '$sections.sectionTitle',
            sectionIcon:  '$sections.sectionIcon',
            sectionType:  '$sections.sectionType',
            itemText:     '$sections.rows.itemText',
            status:       '$sections.rows.status',
          },
          count: { $sum: 1 },
        },
      },
      {
        $group: {
          _id: {
            sectionTitle: '$_id.sectionTitle',
            sectionIcon:  '$_id.sectionIcon',
            sectionType:  '$_id.sectionType',
            itemText:     '$_id.itemText',
          },
          passCount:      { $sum: { $cond: [{ $eq: ['$_id.status', 'Pass']      }, '$count', 0] } },
          failCount:      { $sum: { $cond: [{ $eq: ['$_id.status', 'Fail']      }, '$count', 0] } },
          partialCount:   { $sum: { $cond: [{ $eq: ['$_id.status', 'Partial']   }, '$count', 0] } },
          excellentCount: { $sum: { $cond: [{ $eq: ['$_id.status', 'Excellent'] }, '$count', 0] } },
          goodCount:      { $sum: { $cond: [{ $eq: ['$_id.status', 'Good']      }, '$count', 0] } },
          averageCount:   { $sum: { $cond: [{ $eq: ['$_id.status', 'Average']   }, '$count', 0] } },
          poorCount:      { $sum: { $cond: [{ $eq: ['$_id.status', 'Poor']      }, '$count', 0] } },
          totalCount:     { $sum: '$count' },
        },
      },
      {
        $group: {
          _id: {
            sectionTitle: '$_id.sectionTitle',
            sectionIcon:  '$_id.sectionIcon',
            sectionType:  '$_id.sectionType',
          },
          items: {
            $push: {
              itemText:      '$_id.itemText',
              passCount:     '$passCount',
              failCount:     '$failCount',
              partialCount:  '$partialCount',
              excellentCount:'$excellentCount',
              goodCount:     '$goodCount',
              averageCount:  '$averageCount',
              poorCount:     '$poorCount',
              totalCount:    '$totalCount',
            },
          },
          totalPass:    { $sum: '$passCount'  },
          totalFail:    { $sum: '$failCount'  },
          totalPartial: { $sum: '$partialCount' },
        },
      },
      {
        $project: {
          _id: 0,
          sectionTitle: '$_id.sectionTitle',
          sectionIcon:  '$_id.sectionIcon',
          sectionType:  '$_id.sectionType',
          totalPass:    1,
          totalFail:    1,
          totalPartial: 1,
          items:        1,
        },
      },
      { $sort: { sectionTitle: 1 } },
    ]);

    return res.json({ period, sections: result });
  } catch (err) {
    console.error('Section items error:', err);
    return res.status(500).json({ message: 'Server error.' });
  }
};