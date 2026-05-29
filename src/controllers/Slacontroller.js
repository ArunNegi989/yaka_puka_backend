import SlaCategory from "../models/Slacategorymodel.js";

/* ════════════════════════════════════════════
   HELPER
════════════════════════════════════════════ */
function sendError(res, status, message, details = null) {
  const payload = { success: false, message };
  if (details) payload.details = details;
  return res.status(status).json(payload);
}

/* ════════════════════════════════════════════
   GET /sla/categories
   Returns all categories sorted by order
════════════════════════════════════════════ */
export async function getAllCategories(req, res) {
  try {
    const categories = await SlaCategory.find({}).sort({ order: 1 }).lean();
    return res.status(200).json(categories);
  } catch (err) {
    console.error("[SLA] getAllCategories error:", err);
    return sendError(res, 500, "Failed to fetch SLA categories");
  }
}

/* ════════════════════════════════════════════
   GET /sla/categories/:id
   Returns single category by its frontend id
════════════════════════════════════════════ */
export async function getCategoryById(req, res) {
  try {
    const category = await SlaCategory.findOne({ id: req.params.id }).lean();
    if (!category) {
      return sendError(
        res,
        404,
        `Category with id '${req.params.id}' not found`,
      );
    }
    return res.status(200).json(category);
  } catch (err) {
    console.error("[SLA] getCategoryById error:", err);
    return sendError(res, 500, "Failed to fetch category");
  }
}

/* ════════════════════════════════════════════
   POST /sla/categories
   Create a single new category
════════════════════════════════════════════ */
export async function createCategory(req, res) {
  try {
    const body = req.body;

    if (!body.id) {
      return sendError(res, 400, "Field 'id' is required");
    }
    if (!body.title?.trim()) {
      return sendError(res, 400, "Field 'title' is required");
    }

    // Check duplicate id
    const exists = await SlaCategory.findOne({ id: body.id });
    if (exists) {
      return sendError(
        res,
        409,
        `Category with id '${body.id}' already exists`,
      );
    }

    const category = await SlaCategory.create(body);
    return res.status(201).json({ success: true, data: category });
  } catch (err) {
    console.error("[SLA] createCategory error:", err);
    if (err.name === "ValidationError") {
      return sendError(res, 400, "Validation failed", err.message);
    }
    return sendError(res, 500, "Failed to create category");
  }
}

/* ════════════════════════════════════════════
   PUT /sla/categories
   BULK REPLACE — replaces the entire config.
   This is what the frontend calls on "Save Changes".
   It deletes all existing docs and inserts the new payload.
════════════════════════════════════════════ */
export async function bulkReplaceCategories(req, res) {
  try {
    const payload = req.body;

    // ── Validate payload is an array ──
    if (!Array.isArray(payload)) {
      return sendError(res, 400, "Request body must be an array of categories");
    }

    // ── Per-category validation ──
    for (let i = 0; i < payload.length; i++) {
      const cat = payload[i];
      if (!cat.id) {
        return sendError(res, 400, `Category at index ${i} is missing 'id'`);
      }
      if (!cat.title?.trim()) {
        return sendError(
          res,
          400,
          `Category at index ${i} (id: ${cat.id}) is missing 'title'`,
        );
      }
      if (!["checklist", "feedback"].includes(cat.type)) {
        return sendError(
          res,
          400,
          `Category '${cat.title}' has invalid type '${cat.type}'. Must be 'checklist' or 'feedback'`,
        );
      }

      if (cat.questions && Array.isArray(cat.questions)) {
        for (let j = 0; j < cat.questions.length; j++) {
          const q = cat.questions[j];
          if (!q.text?.trim()) {
            return sendError(
              res,
              400,
              `Category '${cat.title}': question at index ${j} has empty text`,
            );
          }
        }
      }
    }

    // ── Atomic replace ──
    // Delete all → insert new (wrapped in one logical operation)
    await SlaCategory.deleteMany({});

    let inserted = [];
    if (payload.length > 0) {
      inserted = await SlaCategory.insertMany(
        payload.map((cat, i) => ({ ...cat, order: i })),
        { ordered: true },
      );
    }

    return res.status(200).json({
      success: true,
      message: `Saved ${inserted.length} categories`,
      count: inserted.length,
    });
  } catch (err) {
    console.error("[SLA] bulkReplaceCategories error:", err);
    if (err.name === "ValidationError") {
      return sendError(res, 400, "Validation failed", err.message);
    }
    return sendError(res, 500, "Failed to save categories");
  }
}

/* ════════════════════════════════════════════
   PATCH /sla/categories/:id
   Partial update of a single category
════════════════════════════════════════════ */
export async function updateCategory(req, res) {
  try {
    const updated = await SlaCategory.findOneAndUpdate(
      { id: req.params.id },
      { $set: req.body },
      { new: true, runValidators: true },
    );

    if (!updated) {
      return sendError(res, 404, `Category '${req.params.id}' not found`);
    }

    return res.status(200).json({ success: true, data: updated });
  } catch (err) {
    console.error("[SLA] updateCategory error:", err);
    if (err.name === "ValidationError") {
      return sendError(res, 400, "Validation failed", err.message);
    }
    return sendError(res, 500, "Failed to update category");
  }
}

/* ════════════════════════════════════════════
   DELETE /sla/categories/:id
   Delete single category by frontend id
════════════════════════════════════════════ */
export async function deleteCategory(req, res) {
  try {
    const deleted = await SlaCategory.findOneAndDelete({ id: req.params.id });

    if (!deleted) {
      return sendError(res, 404, `Category '${req.params.id}' not found`);
    }

    return res.status(200).json({
      success: true,
      message: `Category '${deleted.title}' deleted`,
    });
  } catch (err) {
    console.error("[SLA] deleteCategory error:", err);
    return sendError(res, 500, "Failed to delete category");
  }
}

/* ════════════════════════════════════════════
   POST /sla/categories/:id/questions
   Add a question to an existing category
════════════════════════════════════════════ */
export async function addQuestion(req, res) {
  try {
    const q = req.body;

    if (!q.id) return sendError(res, 400, "Question 'id' is required");
    if (!q.text?.trim())
      return sendError(res, 400, "Question 'text' is required");

    const category = await SlaCategory.findOne({ id: req.params.id });
    if (!category) {
      return sendError(res, 404, `Category '${req.params.id}' not found`);
    }

    q.order = category.questions.length;
    category.questions.push(q);
    await category.save();

    return res.status(201).json({ success: true, data: category });
  } catch (err) {
    console.error("[SLA] addQuestion error:", err);
    return sendError(res, 500, "Failed to add question");
  }
}

/* ════════════════════════════════════════════
   DELETE /sla/categories/:id/questions/:qid
   Remove a specific question from a category
════════════════════════════════════════════ */
export async function deleteQuestion(req, res) {
  try {
    const category = await SlaCategory.findOne({ id: req.params.id });
    if (!category) {
      return sendError(res, 404, `Category '${req.params.id}' not found`);
    }

    const before = category.questions.length;
    category.questions = category.questions.filter(
      (q) => q.id !== req.params.qid,
    );

    if (category.questions.length === before) {
      return sendError(res, 404, `Question '${req.params.qid}' not found`);
    }

    // Re-order
    category.questions = category.questions.map((q, i) => ({
      ...q.toObject(),
      order: i,
    }));
    await category.save();

    return res.status(200).json({ success: true, data: category });
  } catch (err) {
    console.error("[SLA] deleteQuestion error:", err);
    return sendError(res, 500, "Failed to delete question");
  }
}
