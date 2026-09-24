const interview = require("../routes/interview");
const express = require("express");
const multer = require("multer");


const router = express.Router();

const excelUpload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024,
  },
}).any();

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
router.post("/excelimportSep", excelUpload, interview.excelimportSep);
router.post("/importformatnewjoining",interview.importformatnewjoining)
router.post("/detailedreport",interview.detailedreport)


// ── CEO Dashboard API (Reference & Compatibility) ──
router.post("/getDashboardData", interview.getDashboardData);

// ── Recruitment Process Dashboard API ──
router.post("/candidateDashboard", interview.CandidateDashboard);
router.post("/CandidateDashboard", interview.CandidateDashboard);
router.post("/recruitmentDashboard", interview.CandidateDashboard);

module.exports = router;