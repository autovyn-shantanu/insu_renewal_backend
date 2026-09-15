const interview = require("../routes/interview");
const express = require("express");

const router = express.Router();

// ── Manager Dashboard APIs ──
router.post("/getManagerDashboardData", interview.getManagerDashboardData);
router.post("/getManagerAttendanceStats", interview.getManagerAttendanceStats);
router.post("/getManagerSalaryStats", interview.getManagerSalaryStats);
router.post("/getManagerActionItems", interview.getManagerActionItems);
router.post("/getManagerWatchlist", interview.getManagerWatchlist);
router.post("/getNeedsYouToday", interview.getNeedsYouToday);
router.post("/needsyoutoday", interview.getNeedsYouToday);
router.post("/gethrentry", interview.gethrentry);
router.post("/interviewcanidates", interview.interviewcanidates);
router.post("/shortlistcandidate", interview.shortlistcandidate);


// ── CEO Dashboard API (Reference & Compatibility) ──
router.post("/getDashboardData", interview.getDashboardData);

module.exports = router;