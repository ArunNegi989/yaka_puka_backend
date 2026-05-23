import express from "express";
import {
  getAllCategories,
  getCategoryById,
  createCategory,
  bulkReplaceCategories,
  updateCategory,
  deleteCategory,
  addQuestion,
  deleteQuestion,
} from "../controllers/Slacontroller.js";

const router = express.Router();

/* ════════════════════════════════════════════
   CATEGORY ROUTES
════════════════════════════════════════════ */

// GET    /sla/categories          → get all categories (sorted by order)
router.get("/categories", getAllCategories);

// GET    /sla/categories/:id      → get one category
router.get("/categories/:id", getCategoryById);

// POST   /sla/categories          → create one new category
router.post("/categories", createCategory);

// PUT    /sla/categories          → BULK REPLACE entire config (used by "Save Changes")
router.put("/categories", bulkReplaceCategories);

// PATCH  /sla/categories/:id      → partial update one category
router.patch("/categories/:id", updateCategory);

// DELETE /sla/categories/:id      → delete one category
router.delete("/categories/:id", deleteCategory);

/* ════════════════════════════════════════════
   QUESTION ROUTES (nested under category)
════════════════════════════════════════════ */

// POST   /sla/categories/:id/questions          → add question to category
router.post("/categories/:id/questions", addQuestion);

// DELETE /sla/categories/:id/questions/:qid     → remove question from category
router.delete("/categories/:id/questions/:qid", deleteQuestion);

export default router;