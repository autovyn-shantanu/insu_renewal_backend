const hr = require("../routes/hr");
const express = require("express");

const router = express.Router();

// ── Manager Dashboard APIs ──
router.post("/getManagerDashboardData", hr.getManagerDashboardData);
router.post("/getManagerAttendanceStats", hr.getManagerAttendanceStats);
router.post("/getManagerSalaryStats", hr.getManagerSalaryStats);
router.post("/getManagerActionItems", hr.getManagerActionItems);
router.post("/getManagerWatchlist", hr.getManagerWatchlist);
router.post("/getNeedsYouToday", hr.getNeedsYouToday);
router.post("/needsyoutoday", hr.getNeedsYouToday);

// ── CEO Dashboard API (Reference & Compatibility) ──
router.post("/getDashboardData", hr.getDashboardData);

module.exports = router;