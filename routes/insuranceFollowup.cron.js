const cron = require("node-cron");
const { dbname } = require("../utils/dbconfig");
const { QueryTypes } = require("sequelize");
const { triggerSingleCall, getCallStatus } = require("./callmatics");
const { SendInsuranceRenewalWhatsAppToCustomer, saveCallWebhookDetails } = require("./Crm");
const SchedulerLogger = require("../utils/schedulerLogger");
const fs = require("fs");
const path = require("path");

const CRON_CONFIG = {
  SCHEDULE: "* * * * *", // Har 1 minute me chalega
  TIMEZONE: "Asia/Kolkata",
  CALL_DELAY_MS: 1500,
};

// ============================================================
// MAIN SCHEDULER ASYNC FUNCTION (All functions inside)
// ============================================================
async function runInsuranceFollowupScheduler(targetCompcode = null) {
  console.log("Running scheduled job for Insurance Followup Calling...");

  // ─────────────────────────────────────────────
  // ALL HELPER FUNCTIONS DEFINED INSIDE ASYNC
  // ─────────────────────────────────────────────
  function writeDailySchedulerLog(logData) {
    try {
      const today = new Date().toISOString().split("T")[0];
      const logDir = path.join(__dirname, "../scheduler_logs");

      if (!fs.existsSync(logDir)) {
        fs.mkdirSync(logDir, { recursive: true });
      }

      const filePath = path.join(logDir, `${today}.txt`);
      fs.appendFileSync(filePath, JSON.stringify(logData) + "\n");
    } catch (err) {
      console.error("TXT log failed:", err.message);
    }
  }

  const normalizeVehNo = (v) =>
    String(v || "")
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, "")
      .trim();

  const normalizePhone = (p) => {
    if (!p) return "";
    let s = String(p).trim().replace(/[\s-]/g, "");
    if (/^\d{10}$/.test(s)) s = `+91${s}`;
    if (/^91\d{10}$/.test(s)) s = `+${s}`;
    return s;
  };

  const formatDateYYYYMMDD_to_DDMMYYYY = (dateStr) => {
    if (!dateStr) return "";
    const parts = String(dateStr).split("-");
    if (parts.length !== 3) return String(dateStr);
    const [y, m, d] = parts;
    return `${String(d).padStart(2, "0")}/${String(m).padStart(2, "0")}/${y}`;
  };

  const timeToAmPm = (hhmmss) => {
    if (!hhmmss) return null;
    const s = String(hhmmss).trim();
    const match = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!match) return s;
    let h = parseInt(match[1], 10);
    const min = match[2];
    const ampm = h >= 12 ? "PM" : "AM";
    h = h % 12;
    if (h === 0) h = 12;
    return `${h}:${min} ${ampm}`;
  };

  const toMDY = (d = new Date()) => {
    const mm = d.getMonth() + 1;
    const dd = d.getDate();
    const yyyy = d.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  };

  let sequelize1;
  try {
    sequelize1 = await dbname(
      { query: "", headers: { compcode: "DBCON", name: "schedualer" } },
      "DBCON"
    );
  } catch (e) {
    console.error("[INS-CRON] DBCON connection error:", e?.message);
  }

  // 1. Determine active dealers from DLR_SCH
  let Dlr_data = [];

  if (targetCompcode) {
    Dlr_data = [{ Dlr_Id: String(targetCompcode).trim().toLowerCase() }];
  } else if (sequelize1) {
    try {
      const [rows] = await sequelize1.query(
        `SELECT DISTINCT LTRIM(RTRIM(Dlr_Id)) AS Dlr_Id
         FROM DLR_SCH
         WHERE SCH_TYPE = 'insuranceremider' 
           AND export_type < 3
           AND Dlr_Id IS NOT NULL
           AND LTRIM(RTRIM(Dlr_Id)) <> ''`
      );
      Dlr_data = rows || [];
    } catch (e) {
      console.error("[INS-CRON] DBCON DLR_SCH query error:", e?.message);
      Dlr_data = [];
    }
  }

  console.log(`[INS-CRON] Active dealers to process: ${Dlr_data.length}`);

  const results = [];

  // 2. Loop through each dealer
  for (const dealer of Dlr_data) {
    const compcode = dealer.Dlr_Id;
    if (!compcode) continue;

    const logger = new SchedulerLogger(sequelize1, CRON_CONFIG.SCH_TYPE, compcode);
    console.log(`\n⏳ Processing dealer: ${compcode}`);

    let sequelize;
    try {
      sequelize = await dbname(
        { query: "", headers: { compcode, name: "schedualer" } },
        compcode
      );

      if (!sequelize) {
        console.log(`⚠️ [${compcode}] Skipped: no database connection configured.`);
        logger.addSkip(1);
        continue;
      }

      // Fetch due followups
      const now = new Date();
      const pad2 = (n) => String(n).padStart(2, "0");
      const currDate = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
      const currTime = `${pad2(now.getHours())}:${pad2(now.getMinutes())}:${pad2(now.getSeconds())}`;

      const rows = await sequelize.query(
        `
        SELECT
          f.UTD AS FOLLOWUP_UTD,
          f.TRAN_ID,
          f.FOLLOWUP_STATUS,
          CONVERT(varchar(10), f.FOLLOWUP_DATE, 23) AS FOLLOWUP_DATE,
          CONVERT(varchar(8),  f.FOLLOWUP_TIME, 108) AS FOLLOWUP_TIME,
          f.CALL_ID,

          m.UTD        AS MST_UTD,
          m.VEHICAL_REG_NO,
          m.EXPORT_TYPE,
          m.LOC_CODE,

          r.UTD        AS RENEWAL_UTD,
          r.CUST_NAME,
          r.CUST_MOB_NO,
          r.POLICY_NAME,
          CAST(r.POLICY_NUMBER AS varchar(100)) AS POLICY_NUMBER,
          r.MODEL_NAME,
          r.DSC_MOB_NO,
          CONVERT(varchar(10), r.POLICY_END_DATE, 23) AS POLICY_END_DATE

        FROM dbo.FOLLOWUP_DETAILS f
        INNER JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = f.TRAN_ID

        OUTER APPLY (
          SELECT TOP 1
            rr.UTD, rr.CUST_NAME, rr.CUST_MOB_NO, rr.POLICY_NAME, rr.POLICY_NUMBER,
            rr.MODEL_NAME, rr.POLICY_END_DATE, rr.DSC_MOB_NO
          FROM dbo.INSU_RENEWAL rr
          WHERE rr.TRAN_ID = m.UTD
             OR (
               m.VEHICAL_REG_NO IS NOT NULL 
               AND REPLACE(UPPER(rr.VEHICAL_REG_NO), ' ', '') = REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '')
             )
          ORDER BY rr.UTD DESC
        ) r

        WHERE (m.EXPORT_TYPE = 1 OR m.EXPORT_TYPE IS NULL)
          AND (f.CALL_ID IS NULL OR LTRIM(RTRIM(f.CALL_ID)) = '')
          AND (
            f.FOLLOWUP_STATUS IS NULL
            OR LTRIM(RTRIM(UPPER(f.FOLLOWUP_STATUS))) IN ('PENDING','OPEN','PROMISED','CALLBACK')
          )
          AND r.CUST_MOB_NO IS NOT NULL
          AND LTRIM(RTRIM(CAST(r.CUST_MOB_NO AS varchar(20)))) <> ''

          -- DUE CONDITION (compared with current system date & time)
          AND (
            CAST(f.FOLLOWUP_DATE AS date) < CAST(:currDate AS date)
            OR (
              CAST(f.FOLLOWUP_DATE AS date) = CAST(:currDate AS date)
              AND (
                f.FOLLOWUP_TIME IS NULL
                OR LTRIM(RTRIM(f.FOLLOWUP_TIME)) = ''
                OR TRY_CONVERT(time(0), f.FOLLOWUP_TIME) IS NULL
                OR TRY_CONVERT(time(0), f.FOLLOWUP_TIME) <= CAST(:currTime AS time(0))
              )
            )
          )
        ORDER BY f.UTD ASC
        `,
        {
          replacements: { currDate, currTime },
          type: QueryTypes.SELECT,
        }
      );

      if (!rows || !rows.length) {
        console.log(`✅ No due insurance followups for ${compcode}.`);
        logger.addSkip(1);
        continue;
      }

      console.log(`📞 Found ${rows.length} due followups for ${compcode}`);

      for (const r of rows) {
        try {
          const phoneNumber = normalizePhone(r.CUST_MOB_NO);
          if (!phoneNumber)
            throw new Error(`Invalid customer mobile for TRAN_ID=${r.TRAN_ID}`);

          // Fetch INSU_CALLING_CONFIG dynamically by LOC_CODE
          const locCode = r.LOC_CODE != null ? Number(r.LOC_CODE) : null;
          let cfg = null;
          try {
            const cfgRows = await sequelize.query(
              `SELECT TOP 1
                 C.UTD,
                 C.INSU_COMPANY_NAME,
                 C.SALES_EXECUTIVE_NO,
                 C.SLOT1,
                 C.SLOT2,
                 C.SLOT3,
                 C.CAMPAIGN_ID,
                 CONVERT(varchar(8), C.CALLBACK_TIME) AS CALLBACK_TIME,
                 C.LOC_CODE,
                 C.STATUS
               FROM dbo.INSU_CALLING_CONFIG C
               WHERE (C.LOC_CODE = :locCode OR C.LOC_CODE IS NULL)
               ORDER BY
                 CASE WHEN C.LOC_CODE = :locCode THEN 0 ELSE 1 END,
                 C.UTD DESC`,
              {
                replacements: { locCode: locCode != null ? Number(locCode) : null },
                type: QueryTypes.SELECT,
              }
            );
            cfg = cfgRows?.[0] || null;
          } catch (cfgErr) {
            console.error("[INS-CRON] getCallingConfig error:", cfgErr?.message);
          }

          // Build dynamic variables
          const resolvedVehicleModel = (() => {
            if (r.MODEL_NAME != null && String(r.MODEL_NAME).trim() !== "")
              return String(r.MODEL_NAME).trim();
            if (r.POLICY_NAME != null && String(r.POLICY_NAME).trim() !== "")
              return String(r.POLICY_NAME).trim();
            if (r.VEHICAL_REG_NO != null && String(r.VEHICAL_REG_NO).trim() !== "")
              return normalizeVehNo(r.VEHICAL_REG_NO);
            return null;
          })();

          const resolvedTransferNumber = (() => {
            if (r.DSC_MOB_NO != null && String(r.DSC_MOB_NO).trim() !== "") {
              return String(r.DSC_MOB_NO).trim();
            }
            if (cfg?.SALES_EXECUTIVE_NO != null && String(cfg.SALES_EXECUTIVE_NO).trim() !== "") {
              return String(cfg.SALES_EXECUTIVE_NO).trim();
            }
            return null;
          })();

          const callbackDate = r.FOLLOWUP_DATE
            ? formatDateYYYYMMDD_to_DDMMYYYY(r.FOLLOWUP_DATE)
            : toMDY(new Date());

          const callbackTime = (() => {
            if (r.FOLLOWUP_TIME && String(r.FOLLOWUP_TIME).trim() !== "") {
              return timeToAmPm(r.FOLLOWUP_TIME);
            }
            if (cfg?.CALLBACK_TIME) {
              return timeToAmPm(cfg.CALLBACK_TIME);
            }
            return "04:00 PM";
          })();

          const variables = {
            showroom_name:
              r.POLICY_NAME?.trim() ||
              cfg?.INSU_COMPANY_NAME?.trim() ||
              null,
            vehicle_model: resolvedVehicleModel,
            vehicle_number:
              r.VEHICAL_REG_NO != null && String(r.VEHICAL_REG_NO).trim() !== ""
                ? normalizeVehNo(r.VEHICAL_REG_NO)
                : null,
            insurance_expiry_date:
              formatDateYYYYMMDD_to_DDMMYYYY(r.POLICY_END_DATE) || null,
            transferNumber: resolvedTransferNumber,
            callback_date: callbackDate,
            callback_time: callbackTime,
            callee_name:
              r.CUST_NAME != null && String(r.CUST_NAME).trim() !== ""
                ? String(r.CUST_NAME).trim()
                : null,
          };

          const finalCampaignId = cfg?.CAMPAIGN_ID || null;

          // Validate variables
          const missing = [];
          if (!variables.showroom_name)
            missing.push("showroom_name (INSU_CALLING_CONFIG.INSU_COMPANY_NAME or POLICY_NAME)");
          if (!variables.vehicle_model)
            missing.push("vehicle_model (MODEL_NAME / POLICY_NAME / VEHICAL_REG_NO)");
          if (!variables.vehicle_number)
            missing.push("vehicle_number (INSU_RENEWAL_MST.VEHICAL_REG_NO)");
          if (!variables.insurance_expiry_date)
            missing.push("insurance_expiry_date (INSU_RENEWAL.POLICY_END_DATE)");
          if (!variables.transferNumber)
            missing.push("transferNumber (DSC_MOB_NO / INSU_CALLING_CONFIG.SALES_EXECUTIVE_NO)");
          if (!variables.callback_time)
            missing.push("callback_time (FOLLOWUP_DETAILS.FOLLOWUP_TIME or INSU_CALLING_CONFIG.CALLBACK_TIME)");
          if (!variables.callee_name)
            missing.push("callee_name (INSU_RENEWAL.CUST_NAME)");
          if (!finalCampaignId)
            missing.push("campaign_id (INSU_CALLING_CONFIG.CAMPAIGN_ID)");

          if (missing.length) {
            throw new Error(
              `Missing dynamic config for FOLLOWUP_UTD=${r.FOLLOWUP_UTD} | LOC_CODE=${locCode} | Missing: ${missing.join(", ")}`
            );
          }

          // Trigger call
          const callResult = await triggerSingleCall(
            phoneNumber,
            variables,
            finalCampaignId
          );

          const callId =
            callResult?.callId || callResult?.calls?.[0]?.callId || null;
          if (!callId)
            throw new Error("Call triggered but callId not received from Callmatic");

          // Update FOLLOWUP_DETAILS with callId
          await sequelize.query(
            `
            UPDATE dbo.FOLLOWUP_DETAILS
            SET
              CALL_ID            = :call_id,
              FOLLOWUP_STATUS    = 'AI_CALL_INITIATED',
              LAST_FOLLOWUP_DATE = CAST(GETDATE() AS date)
            WHERE UTD = :followup_utd
            `,
            {
              replacements: {
                call_id: String(callId),
                followup_utd: Number(r.FOLLOWUP_UTD),
              },
              type: QueryTypes.UPDATE,
            }
          );

          // Save call webhook log & call_Id_dtl
          try {
            if (callId && saveCallWebhookDetails) {
              await saveCallWebhookDetails(sequelize, callId, {
                status: "INITIATED",
                phoneNumber,
                calleeName: variables.callee_name || null,
                campaignId: finalCampaignId,
                triggeredAt: new Date(),
                startTime: new Date(),
                category: "INSURANCE_FOLLOWUP",
              });
            }

            await sequelize.query(
              `INSERT INTO dbo.call_Id_dtl (mob_no, call_id, call_type)
               VALUES (:mob_no, :call_id, 'INSURANCE_FOLLOWUP')`,
              {
                replacements: {
                  mob_no: phoneNumber,
                  call_id: String(callId),
                },
                type: QueryTypes.INSERT,
              }
            );
          } catch (_) { }

          // Add Impact to SchedulerLogger
          logger.addImpact("FOLLOWUP_DETAILS", r.FOLLOWUP_UTD, 1);

          console.log(
            `✅ Call initiated for ${compcode} | Cust:${r.CUST_NAME} | Phone:${phoneNumber} | CallID:${callId}`
          );

          // Auto WhatsApp on Call Completion
          setImmediate(async () => {
            try {
              console.log(`[INS-CRON-WA] WhatsApp background monitoring started for callId: ${callId}`);

              const checkCallStatus = async (cId, maxAttempts = 60) => {
                let attempts = 0;
                while (attempts < maxAttempts) {
                  try {
                    const statusResult = await getCallStatus(cId);

                    if (statusResult && sequelize && saveCallWebhookDetails) {
                      await saveCallWebhookDetails(sequelize, cId, {
                        ...statusResult,
                        phoneNumber,
                        calleeName: variables.callee_name || null,
                        campaignId: finalCampaignId,
                      });
                    }

                    const st = String(statusResult?.status || "").toUpperCase();
                    if (["COMPLETED", "ENDED", "FINISHED", "FAILED", "BUSY", "NO_ANSWER", "CANCELED"].includes(st)) {
                      return { success: true, status: statusResult?.status, duration: statusResult?.duration };
                    }
                    await new Promise((resolve) => setTimeout(resolve, 2000));
                    attempts++;
                  } catch (checkErr) {
                    attempts++;
                    await new Promise((resolve) => setTimeout(resolve, 2000));
                  }
                }
                return { success: false, status: "TIMEOUT" };
              };

              if (callId) {
                const callStatusResult = await checkCallStatus(callId);
                if (!callStatusResult?.success) {
                  await new Promise((resolve) => setTimeout(resolve, 5000));
                }
              }

              const effectiveCompCode = compcode || "DBCON";
              const mstUtd = r.MST_UTD || r.TRAN_ID;
              const reqObj = {
                headers: { compcode: effectiveCompCode },
                body: { tranIds: [Number(mstUtd)] },
              };
              let waResult;
              const resObj = {
                status: (code) => ({
                  send: (data) => { waResult = { statusCode: code, data }; return waResult; },
                  json: (data) => { waResult = { statusCode: code, data }; return waResult; },
                }),
              };

              await SendInsuranceRenewalWhatsAppToCustomer(reqObj, resObj);
              console.log(`[INS-CRON-WA] ✅ SendInsuranceRenewalWhatsAppToCustomer completed for MST_UTD=${mstUtd} (${effectiveCompCode})`);
            } catch (waErr) {
              console.error("[INS-CRON-WA] ❌ Auto-WhatsApp failed:", waErr?.message);
            }
          });

          // Delay between calls
          await new Promise((x) => setTimeout(x, CRON_CONFIG.CALL_DELAY_MS));
        } catch (callErr) {
          logger.setError(callErr);
          console.error(`Call failed for FOLLOWUP_UTD=${r.FOLLOWUP_UTD} (${compcode}):`, callErr.message);
        }
      }
    } catch (err) {
      logger.setError(err);
      console.error(`Insurance Followup job failed for ${compcode}:`, err.message);
    } finally {
      if (sequelize) {
        try {
          await sequelize.close();
        } catch (_) { }
      }
    }

    try {
      await logger.save();
      writeDailySchedulerLog(logger.toJSON());
    } catch (logErr) {
      console.error("Logger failed:", logErr.message);
    }

    results.push(logger.toJSON());
  }

  // Close DBCON connection
  if (sequelize1) {
    try {
      await sequelize1.close();
    } catch (_) { }
  }

  console.log(`[INS-CRON] ══ SCHEDULER RUN COMPLETED ══`);
  return targetCompcode ? (results[0] || null) : results;
}

// ─────────────────────────────────────────────
// SINGLE DEALER RUN WRAPPER
// ─────────────────────────────────────────────
// const processDealer = async (compcode) => {
//   return await runInsuranceFollowupScheduler(compcode);
// };

// ─────────────────────────────────────────────
// CRON SCHEDULE REGISTRATION
// ─────────────────────────────────────────────
const startInsuranceFollowupCron = () => {
  cron.schedule(
    CRON_CONFIG.SCHEDULE,
    async () => {
      console.log(
        `\n[INS-CRON] ══ RUN ${new Date().toLocaleString("en-IN")} ══`
      );
      try {
        await runInsuranceFollowupScheduler();
      } catch (e) {
        console.error("[INS-CRON] Fatal error:", e?.message);
      }
    },
    { timezone: CRON_CONFIG.TIMEZONE }
  );

  console.log(
    "[INS-CRON] Insurance Followup cron registered ✅",
    CRON_CONFIG.SCHEDULE
  );
};

module.exports = {
  startInsuranceFollowupCron,
  runInsuranceFollowupScheduler,
  InsuranceFollowupScheduler: runInsuranceFollowupScheduler,

};
