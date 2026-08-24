const express = require("express");
const router = express.Router();

const {
  runInsuranceFollowupScheduler,
  processDealer,
} = require("../routes/insuranceFollowup.cron"); // <-- apna correct path

// OPTIONAL security (recommended)
// const CRON_RUN_KEY = process.env.CRON_RUN_KEY;

router.post("/insurance-followup/run-now", async (req, res) => {
  try {
    const { compcode } = req.body || {};
    if (compcode) {
      const result = await processDealer(String(compcode).trim());
      return res.json({ ok: true, mode: "ONE_DEALER", result });
    }
    const result = await runInsuranceFollowupScheduler();
    return res.json({ ok: true, mode: "ALL_DEALERS", result });
  } catch (e) {
    console.error("RUN-NOW ERROR =>", e); // ✅ terminal me exact error
    return res.status(500).json({
      ok: false,
      message: e?.message || String(e) || "Unknown error",
      stack: e?.stack || null
    });
  }
});

module.exports = router;