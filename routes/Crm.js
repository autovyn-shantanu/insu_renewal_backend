const ExcelJS = require("exceljs");
const { DataTypes, Op } = require("sequelize");
const { dbname } = require("../utils/dbconfig");
const xlsx = require("xlsx");
const { Insu_Renewal } = require("../models/excelImport");
const Sequelize = require("sequelize");
const { Insu_Renewal_Mst } = require("../models/InsuRenewalMst");
const FollowupDetailModel = require("../models/Followup_Detail");
const axios = require("axios");
const { Insu_Renewal_Pymt } = require("../models/Insu_Renewal_Pymt");
const FormData = require("form-data");
const {
  verifyBonvoiceWebhook,
  getLead,
  createLeadAndCall,
  createLeadThenTriggerCall,
} = require("./bonvoice");
const jwt = require("jsonwebtoken");
const {
  triggerSingleCall,
  getCallStatus1,
  getCallRecording1,
} = require("./callmatics");
const { QueryTypes } = require("sequelize");
// const Campain_ID = require("./callmatics");
const { Insu_Renewal_Raw_Data } = require("../models/Insu_Renewal_Raw_Data")
const { SendWhatsAppMessgae } = require("./user");

exports.downloadInsuranceRenualSampleExcel = async (req, res) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("INSURANCE_RENUAL");

  // Freeze header row
  ws.views = [{ state: "frozen", ySplit: 1 }];

  // Only headers (NO ROWS)
  ws.columns = [
    { header: "CUST NAME", key: "CUST_NAME", width: 22 },
    { header: "CUST MOB NO", key: "CUST_MOB_NO", width: 16 },
    { header: "POLICY NAME", key: "POLICY_NAME", width: 22 },
    { header: "POLICY NUMBER", key: "POLICY_NUMBER", width: 18 },
    { header: "VEHICLE REG NO", key: "VEHICAL_REG_NO", width: 18 },
    { header: "MODEL NAME", key: "MODEL_NAME", width: 18 },
    { header: "DSE EMPCODE", key: "DSC_EMPCODE", width: 18 },
    { header: "DSE NAME", key: "DSC_NAME", width: 22 },
    { header: "DSE MOB NO", key: "DSC_MOB_NO", width: 16 },
    { header: "POLICY START DATE", key: "POLICY_START_DATE", width: 20 },
    { header: "POLICY END DATE", key: "POLICY_END_DATE", width: 20 },

  ];

  // Header styling (dark green background + white bold text)
  const headerRow = ws.getRow(1);
  headerRow.height = 20;

  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF0B6A45" }, // dark green
    };
    cell.border = {
      top: { style: "thin", color: { argb: "FFBFBFBF" } },
      left: { style: "thin", color: { argb: "FFBFBFBF" } },
      bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
      right: { style: "thin", color: { argb: "FFBFBFBF" } },
    };
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader(
    "Content-Disposition",
    'attachment; filename="INSURANCE_TEMPLATE.xlsx"',
  );

  await wb.xlsx.write(res);
  res.end();
};


exports.importInsuRenewalExcel = async function (req, res, next) {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const InsuRenewal = Insu_Renewal(sequelize, DataTypes);
    const InsuRenewalMst = Insu_Renewal_Mst(sequelize, DataTypes);
    const InsuRenewalRawData = Insu_Renewal_Raw_Data(sequelize, DataTypes);

    const excelFile = req.files?.["excel"]?.[0] || req.file;
    if (!excelFile)
      return res.status(400).send({ Message: "No file uploaded" });

    const workbook = xlsx.read(excelFile.buffer, {
      type: "buffer",
      cellDates: false,  // ✅ Changed to false (like Warranty)
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) return res.status(400).send({ Message: "Worksheet not found" });

    const sheetHeaders = (
      xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []
    ).map((h) => String(h || "").trim());

    const expectedHeaders = [
      "CUST NAME",
      "CUST MOB NO",
      "POLICY NAME",
      "POLICY NUMBER",
      "VEHICLE REG NO",
      "MODEL NAME",
      "DSE EMPCODE",
      "DSE NAME",
      "DSE MOB NO",
      "POLICY START DATE",
      "POLICY END DATE",

    ];

    const isHeadersMatch =
      sheetHeaders.length === expectedHeaders.length &&
      sheetHeaders.every((h, i) => h === expectedHeaders[i]);

    if (!isHeadersMatch) {
      return res.status(400).send({
        Message:
          "Invalid Excel format! Please upload the file with the correct template.",
        ExpectedHeaders: expectedHeaders,
        FoundHeaders: sheetHeaders,
      });
    }

    const rows = xlsx.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false,
      range: 1,
      header: expectedHeaders,
      dateNF: "dd-mm-yyyy",  // ✅ Changed to dd-mm-yyyy (like Warranty)
    });

    if (!rows.length) {
      return res
        .status(400)
        .send({ Message: "No data found in Excel or Invalid format" });
    }

    const locCodeRaw =
      req.body?.branch ??
      req.body?.loc_code ??
      req.body?.Loc_Code ??
      req.headers?.loc_code ??
      req.headers?.branch ??
      null;

    let locCodeToSave = null;
    if (locCodeRaw != null && String(locCodeRaw).trim() !== "") {
      const parts = String(locCodeRaw).split(",").map((p) => p.trim()).filter(Boolean);
      for (const p of parts) {
        const n = Number(p);
        if (Number.isFinite(n) && n > 0) {
          locCodeToSave = n;
          break;
        }
      }
      if (locCodeToSave === null) {
        const n = Number(String(locCodeRaw).trim());
        if (Number.isFinite(n) && n > 0) locCodeToSave = n;
      }
    }

    console.log(
      "[IMPORT] locCodeRaw:",
      locCodeRaw,
      "=> locCodeToSave:",
      locCodeToSave,
    );

    const mstLocAttr = InsuRenewalMst?.rawAttributes?.LOC_CODE
      ? "LOC_CODE"
      : InsuRenewalMst?.rawAttributes?.loc_code
        ? "loc_code"
        : null;

    // =========================================================
    // ✅ DATE HELPERS (SAME AS IMPORTWARRANTY)
    // =========================================================

    function parseDateFromCell(cell) {
      if (!cell || !cell.v) return null;

      if (cell.t === 'n' && typeof cell.v === 'number') {
        // Serial number (e.g., 44994)
        return new Date((cell.v - 25569) * 86400 * 1000);  // UTC date
      } else if (cell.t === 's' && typeof cell.v === 'string') {
        // String format, handle various formats
        const trimmed = cell.v.trim();

        // Try DD/MM/YYYY
        if (trimmed.includes('/')) {
          const parts = trimmed.split('/');
          if (parts.length === 3) {
            const [d, m, y] = parts.map(p => parseInt(p, 10));
            if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
            const year = y < 100 ? 2000 + y : y;  // Handle 2-digit years
            return new Date(Date.UTC(year, m - 1, d));  // UTC date
          }
        }
        // Try DD-MM-YYYY
        else if (trimmed.includes('-')) {
          const parts = trimmed.split('-');
          if (parts.length === 3) {
            // Try DD-MM-YYYY first
            const [d, m, y] = parts.map(p => parseInt(p, 10));
            if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
              const year = y < 100 ? 2000 + y : y;
              return new Date(Date.UTC(year, m - 1, d));
            }
            // Try YYYY-MM-DD
            const [y2, m2, d2] = parts.map(p => parseInt(p, 10));
            if (!isNaN(d2) && !isNaN(m2) && !isNaN(y2)) {
              return new Date(Date.UTC(y2, m2 - 1, d2));
            }
          }
        }
        // Try direct date parsing
        const parsed = new Date(trimmed);
        if (!isNaN(parsed.getTime())) return parsed;
      }
      return null;  // Invalid type
    }

    // ✅ Function to convert date to SQL format (same as Warranty)
    function toSqlDateTime(date) {
      if (!date || isNaN(date.getTime())) return null;
      const d = new Date(date);
      return d.toISOString().slice(0, 19).replace('T', ' ').split(' ')[0];  // Return only YYYY-MM-DD
    }

    // ✅ Function to get cell for date parsing
    function getDateCell(rowIndex, columnKey, rows, sheet) {
      const rowNum = rowIndex + 2; // Excel rows are 1-indexed, header is row 1

      // Find the column index based on header
      const firstRow = rows[0];
      const keys = Object.keys(firstRow);
      const colIndex = keys.findIndex(key => {
        // Normalize key comparison
        const normalizedKey = key.trim();
        const normalizedColumnKey = columnKey.trim();
        return normalizedKey === normalizedColumnKey;
      });

      if (colIndex === -1) {
        console.warn(`Column "${columnKey}" not found in Excel sheet`);
        return null;
      }

      const colLetter = xlsx.utils.encode_col(colIndex);
      const cellRef = `${colLetter}${rowNum}`;
      return sheet[cellRef];
    }

    const clean = (v) => {
      if (v === null || v === undefined || v === "") return null;
      const s = String(v).trim();
      return s === "" ? null : s;
    };

    const cleanMob = (v) => {
      const s = clean(v);
      if (!s) return null;
      const digits = s.replace(/\D/g, "");
      return digits === "" ? null : digits;
    };

    const cleanRegNo = (v) => {
      const s = clean(v);
      if (!s) return null;
      return s.toUpperCase();
    };

    const cleanDSCMob = (v) => {
      const s = clean(v);
      if (!s) return null;
      const digits = s.replace(/\D/g, "");
      if (digits === "") return null;
      return digits;
    };

    // ✅ Build data with date parsing (same as Warranty)
    const data = rows.map((r, index) => {
      // Find date cells (same logic as Warranty)
      const policyStartDateCell = getDateCell(index, "POLICY START DATE", rows, sheet);
      const policyEndDateCell = getDateCell(index, "POLICY END DATE", rows, sheet);

      // Parse dates
      const startDate = parseDateFromCell(policyStartDateCell);
      const endDate = parseDateFromCell(policyEndDateCell);

      // Convert to SQL format
      const start = toSqlDateTime(startDate);

      // Calculate end date if start date exists
      let autoEnd = null;
      if (start) {
        const [year, month, day] = start.split('-').map(Number);
        const nextYear = new Date(Date.UTC(year + 1, month - 1, day));
        autoEnd = nextYear.toISOString().slice(0, 19).replace('T', ' ').split(' ')[0];
      }

      return {
        CUST_NAME: clean(r["CUST NAME"]),
        CUST_MOB_NO: cleanMob(r["CUST MOB NO"]),
        POLICY_NAME: clean(r["POLICY NAME"]),
        POLICY_NUMBER: clean(r["POLICY NUMBER"]),
        VEHICAL_REG_NO: cleanRegNo(r["VEHICLE REG NO"]),
        MODEL_NAME: clean(r["MODEL NAME"]),
        POLICY_START_DATE: start,
        POLICY_END_DATE: autoEnd,
        DSC_EMPCODE: clean(r["DSE EMPCODE"]),
        DSC_NAME: clean(r["DSE NAME"]),
        DSC_MOB_NO: cleanDSCMob(r["DSE MOB NO"]),
      };
    });

    const CorrectData = [];
    const ErroredData = [];
    const RawDataErrors = [];
    let updatedVehicles = 0;
    let rawInserted = 0;

    // =================================================================
    // STEP 1: Save ALL Excel rows AS-IS to INSU_RENEWAL_RAW_DATA
    // =================================================================
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const excelRowNo = i + 2;

      try {
        await InsuRenewalRawData.create(
          {
            CUST_NAME: row.CUST_NAME ?? "",
            CUST_MOB_NO: row.CUST_MOB_NO,
            POLICY_NAME: row.POLICY_NAME ?? "",
            POLICY_NUMBER: row.POLICY_NUMBER ?? "",
            VEHICAL_REG_NO: row.VEHICAL_REG_NO,
            MODEL_NAME: row.MODEL_NAME,
            POLICY_START_DATE: row.POLICY_START_DATE,
            POLICY_END_DATE: row.POLICY_END_DATE,
            DSC_EMPCODE: row.DSC_EMPCODE,
            DSC_NAME: row.DSC_NAME,
            DSC_MOB_NO: row.DSC_MOB_NO,
            CREATED_AT: Sequelize.literal("GETDATE()"),
          },
          {
            validate: false,
            fields: [
              "CUST_NAME",
              "CUST_MOB_NO",
              "POLICY_NAME",
              "POLICY_NUMBER",
              "VEHICAL_REG_NO",
              "MODEL_NAME",
              "POLICY_START_DATE",
              "POLICY_END_DATE",
              "DSC_EMPCODE",
              "DSC_NAME",
              "DSC_MOB_NO",
              "CREATED_AT",
            ],
          },
        );
        rawInserted++;
      } catch (rawErr) {
        console.error(
          `RAW DATA save failed for Excel row ${excelRowNo}:`,
          rawErr.message,
        );
        RawDataErrors.push({ Excel_Row: excelRowNo, error: rawErr.message });
      }
    }

    // =================================================================
    // STEP 2: Validate & Import to INSU_RENEWAL + INSU_RENEWAL_MST
    // =================================================================
    await sequelize.transaction(async (t) => {
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const excelRowNo = i + 2;

        try {
          // VALIDATIONS
          const regNoRaw = row.VEHICAL_REG_NO ? String(row.VEHICAL_REG_NO) : "";
          const regNo = regNoRaw.trim().toUpperCase();

          if (!regNo) throw new Error("VEHICLE REG NO is required");
          if (!/^[A-Z0-9 \-]+$/.test(regNo))
            throw new Error("Invalid VEHICLE REG NO format");



          const mob = row.CUST_MOB_NO ? String(row.CUST_MOB_NO).trim() : "";
          if (!mob) throw new Error("CUST MOB NO is required");
          const mobDigits = mob.replace(/\D/g, "");
          if (mobDigits.length !== 10)
            throw new Error("Invalid CUST MOB NO (must be 10 digits)");

          if (!row.POLICY_START_DATE)
            throw new Error("POLICY START DATE is required / invalid format");

          if (!row.POLICY_END_DATE)
            throw new Error(
              "Unable to calculate POLICY END DATE from POLICY START DATE",
            );

          if (row.DSC_MOB_NO !== null && row.DSC_MOB_NO !== undefined) {
            const DSCMobStr = String(row.DSC_MOB_NO).replace(/\D/g, "");
            if (DSCMobStr.length !== 10)
              throw new Error("Invalid DSC MOB NO (must be 10 digits)");
            row.DSC_MOB_NO = DSCMobStr;
          }

          let tranId = null;

          if (regNo) {
            const mstExisting = await InsuRenewalMst.findOne({
              attributes: ["UTD"],
              where: { VEHICAL_REG_NO: regNo },
              raw: true,
              transaction: t,
            });

            if (mstExisting) {
              tranId = mstExisting.UTD;
              if (locCodeToSave != null) {
                await sequelize.query(
                  `UPDATE dbo.INSU_RENEWAL_MST
                   SET LOC_CODE = :locCode, EXPORT_TYPE = 1
                   WHERE UTD = :utd`,
                  {
                    replacements: { locCode: locCodeToSave, utd: tranId },
                    transaction: t,
                  },
                );
              }
            } else {
              // ✅ APPLY LOC_CODE while creating MST
              const mstCreateObj = {
                VEHICAL_REG_NO: regNo,
                EXPORT_TYPE: 1,
                LOC_CODE: locCodeToSave,
                CREATED_AT: Sequelize.literal("GETDATE()"),
              };

              const mstFields = ["VEHICAL_REG_NO", "EXPORT_TYPE", "LOC_CODE", "CREATED_AT"];

              const mstNew = await InsuRenewalMst.create(mstCreateObj, {
                transaction: t,
                validate: false,
                fields: mstFields,
              });

              tranId = mstNew?.UTD ?? mstNew?.dataValues?.UTD ?? null;

              // ✅ Extra safety: force DB column update
              if (tranId && locCodeToSave != null) {
                await sequelize.query(
                  `UPDATE dbo.INSU_RENEWAL_MST
                   SET LOC_CODE = :locCode, EXPORT_TYPE = 1
                   WHERE UTD = :utd`,
                  {
                    replacements: { locCode: locCodeToSave, utd: tranId },
                    transaction: t,
                  },
                );
              }
            }

            const detailCount = await InsuRenewal.count({
              where: { VEHICAL_REG_NO: regNo },
              transaction: t,
            });

            if (detailCount > 0) {
              await InsuRenewal.update(
                { EXPORT_TYPE: 33 },
                { where: { VEHICAL_REG_NO: regNo }, transaction: t },
              );
              updatedVehicles += 1;
            }
          }

          const insertObj = {
            ...row,
            TRAN_ID: tranId,
            EXPORT_TYPE: 1,
            CREATED_AT: Sequelize.literal("GETDATE()"),
          };

          const created = await InsuRenewal.create(insertObj, {
            transaction: t,
            validate: false,
            fields: [
              "CUST_NAME",
              "CUST_MOB_NO",
              "POLICY_NAME",
              "POLICY_NUMBER",
              "VEHICAL_REG_NO",
              "MODEL_NAME",
              "POLICY_START_DATE",
              "POLICY_END_DATE",
              "CREATED_AT",
              "EXPORT_TYPE",
              "TRAN_ID",
              "DSC_EMPCODE",
              "DSC_NAME",
              "DSC_MOB_NO",
            ],
          });

          if (tranId) {
            const utd = created?.UTD ?? created?.dataValues?.UTD;
            if (utd) {
              await sequelize.query(
                "UPDATE dbo.INSU_RENEWAL SET TRAN_ID = :tranId WHERE UTD = :utd",
                {
                  replacements: { tranId, utd },
                  transaction: t,
                },
              );
            }
          }

          CorrectData.push({ ...insertObj, Excel_Row: excelRowNo });
        } catch (e) {
          ErroredData.push({
            ...row,
            Excel_Row: excelRowNo,
            rejectionReasons: e.message,
          });
        }
      }
    });

    const workbook1 = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(
      workbook1,
      xlsx.utils.json_to_sheet(CorrectData || []),
      "Imported Data",
    );
    xlsx.utils.book_append_sheet(
      workbook1,
      xlsx.utils.json_to_sheet(ErroredData || []),
      "Non Imported Data",
    );

    const buffer = xlsx.write(workbook1, { type: "buffer", bookType: "xlsx" });
    const base64File = Buffer.from(buffer).toString("base64");

    return res.status(200).json({
      Type: CorrectData.length > 0 ? "success" : "warning",
      Inserted: CorrectData.length,
      UpdatedVehiclesTo33: updatedVehicles,
      NonInserted: ErroredData.length,
      Total: data.length,
      RawInserted: rawInserted,
      RawFailed: RawDataErrors.length,
      RawErrors: RawDataErrors,
      Message:
        CorrectData.length > 0
          ? `${CorrectData.length} Records Processed`
          : "Data Is Not Valid To Import",
      File: base64File,
      FileName: "insu_renewal_import_result.xlsx",
    });
  } catch (error) {
    console.error("Error during file import:", error);
    return res.status(500).json({
      Message: "An error occurred during file import.",
      Error: error.message,
    });
  } finally {
    await sequelize.close();
  }
};


exports.getInsuRenewalByDateRange = async function (req, res, next) {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const { QueryTypes } = require("sequelize");

    const pad2 = (x) => String(x).padStart(2, "0");

    // =========================================================
    // Convert incoming date to YYYY-MM-DD
    // Supports:
    // DD/MM/YYYY
    // DD-MM-YYYY
    // YYYY-MM-DD
    // =========================================================
    function toDbDateOnly(v) {
      if (!v) return null;

      const s = String(v).trim();

      if (!s) return null;

      let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

      if (m) {
        return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
      }

      m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);

      if (m) {
        return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
      }

      m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

      if (m) {
        return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
      }

      return null;
    }

    // =========================================================
    // Convert DB date to DD/MM/YYYY
    // =========================================================
    function toDisplayDDMMYYYY(v) {
      if (!v) return null;

      if (typeof v === "string") {
        const m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);

        if (m) {
          return `${m[3]}/${m[2]}/${m[1]}`;
        }
      }

      const d = v instanceof Date ? v : new Date(v);

      if (isNaN(d.getTime())) return null;

      return `${pad2(d.getDate())}/${pad2(
        d.getMonth() + 1
      )}/${d.getFullYear()}`;
    }

    // =========================================================
    // REQUEST BODY
    // =========================================================
    const {
      fromDate,
      toDate,
      empcode,
      emp_dms_code,
    } = req.body;

    // =========================================================
    // EMPLOYEE CODE REQUIRED
    // =========================================================
    if (!empcode) {
      return res.status(400).json({
        ok: false,
        Message: "empcode required",
        example: {
          fromDate: "26/07/2026",
          toDate: "27/07/2026",
          empcode: "19012719",
        },
      });
    }

    // =========================================================
    // DATE VALIDATION
    // =========================================================
    const from = toDbDateOnly(fromDate);
    const to = toDbDateOnly(toDate);

    if (!from || !to) {
      return res.status(400).json({
        ok: false,
        Message: "fromDate & toDate required (DD/MM/YYYY format)",
      });
    }

    // =========================================================
    // HANDLE REVERSE DATE
    // =========================================================
    const fromFinal = from <= to ? from : to;
    const toFinal = from <= to ? to : from;

    // =========================================================
    // STEP 1
    // LOGIN EMPLOYEE KA BRANCH NIKALO
    // =========================================================
    const empBranchSql = `
        SELECT TOP 1
          loc_code
        FROM EMPLOYEEMASTER
        WHERE EMPCODE = :empcode
      `;

    const empBranchData = await sequelize.query(empBranchSql, {
      replacements: {
        empcode,
      },
      type: QueryTypes.SELECT,
    });

    if (!empBranchData || empBranchData.length === 0) {
      return res.status(404).json({
        ok: false,
        Message: "Employee not found",
        empCode: empcode,
      });
    }

    const userBranch = empBranchData[0].loc_code;

    // =========================================================
    // STEP 2
    // LOGIN EMPLOYEE KE UNDER REPORTING EMPLOYEES
    // =========================================================
    const reportingSql = `
        SELECT DISTINCT
          EMPCODE
        FROM EMPLOYEEMASTER
        WHERE
          (
            Reporting_1 = :empcode
            OR Reporting_2 = :empcode
            OR Reporting_3 = :empcode
          )
          AND loc_code = :userBranch
      `;

    const reportingEmps = await sequelize.query(reportingSql, {
      replacements: {
        empcode,
        userBranch,
      },
      type: QueryTypes.SELECT,
    });

    // =========================================================
    // STEP 3
    // EMPLOYEE CODE LIST
    // =========================================================

    const reportingEmpCodes = reportingEmps
      .map((e) => e.EMPCODE)
      .filter((code) => code !== null && code !== undefined)
      .map((code) => String(code).trim())
      .filter((code) => code !== "");

    // Login employee ka khud ka data bhi include karo
    if (!reportingEmpCodes.includes(String(empcode).trim())) {
      reportingEmpCodes.push(String(empcode).trim());
    }

    // Duplicate remove
    const uniqueEmpCodes = [...new Set(reportingEmpCodes)];


    // =========================================================
    // STEP 4
    // CHECK IF emp_dms_code EXISTS (ADMIN CHECK)
    // ✅ अगर emp_dms_code दिया गया है तो सभी data दिखाओ
    // नहीं तो सिर्फ subordinates का data दिखाओ
    // =========================================================

    let whereCondition = "";
    let replacements = {
      fromFinal,
      toFinal,
      userBranch,
    };

    const isAdmin =
      String(emp_dms_code || "").trim().toUpperCase() === "EDP";
    if (isAdmin) {
      console.log("✅ ADMIN MODE: Showing all data");
      // Admin mode: सभी data दिखाओ, कोई employee filter नहीं
      whereCondition = `
          WHERE
            ISNULL(r.EXPORT_TYPE, 1) = 1
            AND ISNULL(m.EXPORT_TYPE, 1) <> 33
            AND TRY_CONVERT(date, r.CREATED_AT)
                BETWEEN :fromFinal AND :toFinal
            AND (m.loc_code = :userBranch OR m.loc_code IS NULL OR m.loc_code = 0 OR :userBranch IS NULL)
        `;
    } else {
      // Normal mode: सिर्फ subordinates का data
      console.log("👤 EMPLOYEE MODE: Showing subordinate data only");

      // =========================================================
      // DYNAMIC PARAMETER PLACEHOLDERS FOR IN CLAUSE
      // =========================================================

      const employeePlaceholders = uniqueEmpCodes
        .map((_, index) => `:emp${index}`)
        .join(", ");

      uniqueEmpCodes.forEach((code, index) => {
        replacements[`emp${index}`] = code;
      });

      whereCondition = `
          WHERE
            ISNULL(r.EXPORT_TYPE, 1) = 1
            AND ISNULL(m.EXPORT_TYPE, 1) <> 33
            AND TRY_CONVERT(date, r.CREATED_AT)
                BETWEEN :fromFinal AND :toFinal
            AND (m.loc_code = :userBranch OR m.loc_code IS NULL OR m.loc_code = 0 OR :userBranch IS NULL)
            AND (r.DSC_EMPCODE IN (${employeePlaceholders}))
        `;
    }

    // =========================================================
    // STEP 5
    // FINAL QUERY
    // =========================================================

    const sql = `
        SELECT
          r.*
        FROM dbo.INSU_RENEWAL r

        INNER JOIN dbo.INSU_RENEWAL_MST m
          ON m.UTD = r.TRAN_ID

        ${whereCondition}

        ORDER BY r.UTD DESC
      `;

    console.log("Final SQL:");
    console.log(sql);

    console.log("SQL replacements:");
    console.log(replacements);

    // =========================================================
    // EXECUTE QUERY
    // =========================================================

    const rows = await sequelize.query(sql, {
      replacements,
      type: QueryTypes.SELECT,
    });

    // =========================================================
    // FORMAT RESPONSE DATA
    // =========================================================

    const data = (rows || []).map((r) => ({
      ...r,

      POLICY_START_DATE: toDisplayDDMMYYYY(
        r.POLICY_START_DATE
      ),

      POLICY_END_DATE: toDisplayDDMMYYYY(
        r.POLICY_END_DATE
      ),

      CREATED_AT: toDisplayDDMMYYYY(
        r.CREATED_AT
      ),

      VALIDFROM: toDisplayDDMMYYYY(
        r.VALIDFROM
      ),

      VALIDTO: toDisplayDDMMYYYY(
        r.VALIDTO
      ),
    }));

    // =========================================================
    // SUCCESS RESPONSE
    // =========================================================

    return res.json({
      ok: true,

      loginEmpCode: empcode,

      isAdmin: emp_dms_code && String(emp_dms_code).trim() !== "" ? true : false,

      userBranch,

      subordinateEmpcodes: emp_dms_code && String(emp_dms_code).trim() !== "" ? null : uniqueEmpCodes,

      fromDate: fromFinal,

      toDate: toFinal,

      count: data.length,

      data,
    });

  } catch (error) {

    // =========================================================
    // DETAILED ERROR LOGGING
    // =========================================================

    console.error("========================================");
    console.error("INSURANCE RENEWAL DATE FILTER ERROR");
    console.error("========================================");

    console.error("Error name:", error?.name);
    console.error("Error message:", error?.message);
    console.error("Error parent:", error?.parent?.message);
    console.error("Error original:", error?.original?.message);

    // Sequelize + Tedious multiple RequestErrors
    if (error?.parent?.errors) {
      console.error("----- SQL SERVER ERRORS -----");

      error.parent.errors.forEach((err, index) => {
        console.error(`SQL ERROR ${index + 1}`);
        console.error("Message:", err?.message);
        console.error("Code:", err?.code);
        console.error("Number:", err?.number);
        console.error("State:", err?.state);
        console.error("Class:", err?.class);
        console.error("--------------------------------");
      });
    }

    console.error("SQL:", error?.sql);
    console.error("========================================");

    return res.status(500).json({
      ok: false,
      Message: error?.message || "Something went wrong",
      sqlError:
        error?.parent?.errors?.map((err) => ({
          message: err?.message,
          code: err?.code,
          number: err?.number,
        })) || [],
    });

  } finally {

    await sequelize.close();

  }
};


exports.getInsuRenewalReminders = async function (req, res, next) {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const InsuRenewal = Insu_Renewal(sequelize, DataTypes);
    const { Op } = require("sequelize");

    const ymd = (d = new Date()) => {
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    };

    const addDaysYMD = (baseYmd, days) => {
      const d = new Date(`${baseYmd}T00:00:00`);
      d.setDate(d.getDate() + Number(days));
      return ymd(d);
    };

    const diffDays = (fromYmd, toYmd) => {
      const a = new Date(`${fromYmd}T00:00:00`);
      const b = new Date(`${toYmd}T00:00:00`);
      return Math.round((b - a) / (1000 * 60 * 60 * 24));
    };

    const parseDays = (value, defaultVal) => {
      if (value === null || value === undefined) return defaultVal;
      const s = String(value).trim();
      if (!s) return defaultVal;
      const n = parseInt(s, 10);
      return Number.isFinite(n) ? n : defaultVal;
    };

    const toDbDateOnly = (v) => {
      if (!v) return null;
      const s = String(v).trim();
      if (!s) return null;

      const pad2 = (x) => String(x).padStart(2, "0");

      let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;

      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

      return null;
    };

    const ensureYmdOnly = (v) => {
      if (!v) return null;
      if (v instanceof Date) return ymd(v);
      const s = String(v);
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      return toDbDateOnly(s);
    };

    // =========================================================
    // REQUEST BODY
    // =========================================================
    const { fromDate, toDate, loc_code, empcode, emp_dms_code } = req.body;

    // =========================================================
    // EMPLOYEE CODE REQUIRED
    // =========================================================
    if (!empcode) {
      return res.status(400).json({
        ok: false,
        Message: "empcode required",
      });
    }

    // =========================================================
    // ✅ loc_code normalize
    // =========================================================
    let locCodes = [];
    if (Array.isArray(loc_code)) {
      locCodes = loc_code
        .map((x) => Number(x))
        .filter((n) => Number.isFinite(n));
    } else if (
      loc_code !== undefined &&
      loc_code !== null &&
      String(loc_code).trim() !== ""
    ) {
      const s = String(loc_code);
      if (s.includes(",")) {
        locCodes = s
          .split(",")
          .map((x) => Number(String(x).trim()))
          .filter((n) => Number.isFinite(n));
      } else {
        const n = Number(s);
        if (Number.isFinite(n)) locCodes = [n];
      }
    }

    // 1) Read config
    const [keyRow] = await sequelize.query(
      `SELECT TOP 1 INSU_BEFORE_DAYS, INSU_AFTER_DAYS FROM dbo.COMP_KEYDATA`,
      { type: sequelize.QueryTypes.SELECT }
    );

    const beforeDays = parseDays(keyRow?.INSU_BEFORE_DAYS, 0);
    const afterDays = parseDays(keyRow?.INSU_AFTER_DAYS, 0);

    const today = ymd();
    const beforeTo = addDaysYMD(today, beforeDays);
    const afterFrom = addDaysYMD(today, -afterDays);
    const yesterday = addDaysYMD(today, -1);

    // ✅ EXPORT_TYPE
    const exportType = parseInt(req.body?.EXPORT_TYPE ?? 1, 10);
    const exportTypeFinal = Number.isFinite(exportType) ? exportType : 1;

    // =========================================================
    // ✅ NEW: ADMIN CHECK + EMPLOYEE FILTER
    // =========================================================
    let employeePlaceholders = "";
    let replacements = {
      exportType: exportTypeFinal,
    };

    // अगर emp_dms_code है (Admin) तो सभी data
    // नहीं तो emp_dms_code के subordinates + खुद का data
    const isAdmin =
      String(emp_dms_code || "").trim().toUpperCase() === "EDP";

    if (isAdmin) {
      console.log("✅ ADMIN MODE: Showing all reminders");
      // Admin mode: कोई employee filter नहीं
      employeePlaceholders = "";
    } else {
      console.log("👤 EMPLOYEE MODE: Getting subordinates");
      // Normal employee: subordinates + खुद का data

      // STEP 1: Login employee के subordinates निकालो
      const reportingSql = `
        SELECT DISTINCT EMPCODE
        FROM EMPLOYEEMASTER
        WHERE (
          Reporting_1 = :empcode
          OR Reporting_2 = :empcode
          OR Reporting_3 = :empcode
        )
      `;

      const reportingEmps = await sequelize.query(reportingSql, {
        replacements: { empcode },
        type: sequelize.QueryTypes.SELECT,
      });

      // STEP 2: Employee codes को list में डालो
      let empCodeList = reportingEmps
        .map((e) => e.EMPCODE)
        .filter((code) => code !== null && code !== undefined)
        .map((code) => String(code).trim())
        .filter((code) => code !== "");

      // STEP 3: Login employee को खुद add करो
      if (!empCodeList.includes(String(empcode).trim())) {
        empCodeList.push(String(empcode).trim());
      }

      // STEP 4: Duplicates remove करो
      const uniqueEmpCodes = [...new Set(empCodeList)];

      console.log("Employee codes for filter:", uniqueEmpCodes);

      // STEP 5: Dynamic placeholders बनाओ
      employeePlaceholders = uniqueEmpCodes
        .map((_, index) => `:dscEmp${index}`)
        .join(", ");

      uniqueEmpCodes.forEach((code, index) => {
        replacements[`dscEmp${index}`] = code;
      });
    }

    // ✅ Branch filter -> MST UTD list
    let mstUtdList = null;
    if (locCodes.length) {
      replacements.locCodes = locCodes;
      const mstIds = await sequelize.query(
        `
        SELECT UTD
        FROM dbo.INSU_RENEWAL_MST
        WHERE ISNULL(EXPORT_TYPE, 0) = :exportType
          AND loc_code IN (:locCodes)
        `,
        {
          type: sequelize.QueryTypes.SELECT,
          replacements,
        }
      );

      mstUtdList = (mstIds || [])
        .map((r) => r.UTD)
        .filter(Boolean);

      // no records for that branch => return empty
      if (!mstUtdList.length) {
        return res.json({
          ok: true,
          config: {
            beforeDays,
            afterDays,
            exportType: exportTypeFinal,
            loc_code: locCodes,
            isAdmin: emp_dms_code && String(emp_dms_code).trim() !== "",
          },
          range: {
            today,
            beforeReminderTo: beforeTo,
            afterReminderFrom: afterFrom,
            afterReminderTo: yesterday,
          },
          counts: { before: 0, after: 0 },
          beforeReminders: [],
          afterReminders: [],
          debug: { blockedVehCount: 0, branchFiltered: true },
        });
      }
    }

    // =========================================================
    // ✅ BLOCKED VEH logic (with employee + branch filter)
    // =========================================================
    let blockedVehSQL = `
      ;WITH BlockVeh AS (
        SELECT
          REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '') AS VEH_KEY,
          MAX(CONVERT(date, f.FOLLOWUP_DATE)) AS MAX_FU,
          MAX(CASE 
                WHEN UPPER(LTRIM(RTRIM(f.FOLLOWUP_STATUS))) IN ('NOT_INTERESTED','RENEWED') 
                THEN 1 ELSE 0 
              END) AS IS_CLOSED
        FROM dbo.FOLLOWUP_DETAILS f
        JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = f.TRAN_ID
        WHERE f.FOLLOWUP_DATE IS NOT NULL
          AND m.EXPORT_TYPE = :exportType
    `;

    // ✅ Add employee filter अगर admin नहीं है
    if (employeePlaceholders !== "") {
      blockedVehSQL += `
          AND EXISTS (
            SELECT 1 FROM dbo.INSU_RENEWAL ir
            WHERE ir.TRAN_ID = m.UTD
              AND ir.DSC_EMPCODE IN (${employeePlaceholders})
          )
      `;
    }

    // ✅ Add branch filter
    if (locCodes.length) {
      blockedVehSQL += ` AND m.loc_code IN (:locCodes)`;
    }

    blockedVehSQL += `
        GROUP BY REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '')
      )
      SELECT VEH_KEY
      FROM BlockVeh
      WHERE IS_CLOSED = 1
    `;

    replacements.today = today;

    const blockedVehRows = await sequelize.query(blockedVehSQL, {
      type: sequelize.QueryTypes.SELECT,
      replacements,
    });

    const blockedVehKeys = (blockedVehRows || [])
      .map((r) => r.VEH_KEY)
      .filter(Boolean);

    // =========================================================
    // Build WHERE conditions
    // =========================================================
    const andConds = [{ EXPORT_TYPE: exportTypeFinal }];

    // ✅ Branch filter
    if (mstUtdList && mstUtdList.length) {
      andConds.push({ TRAN_ID: { [Op.in]: mstUtdList } });
    }

    // ✅ Blocked vehicles filter
    if (blockedVehKeys.length) {
      andConds.push(
        sequelize.where(
          sequelize.fn(
            "REPLACE",
            sequelize.fn("UPPER", sequelize.col("VEHICAL_REG_NO")),
            " ",
            ""
          ),
          { [Op.notIn]: blockedVehKeys }
        )
      );
    }

    // ✅ Employee filter (अगर admin नहीं है)
    if (employeePlaceholders !== "") {
      const uniqueEmpCodes = Object.keys(replacements)
        .filter((k) => k.startsWith("dscEmp"))
        .map((k) => replacements[k]);

      andConds.push({ DSC_EMPCODE: { [Op.in]: uniqueEmpCodes } });
    }

    const baseWhere = { [Op.and]: andConds };

    // =========================================================
    // Custom date range
    // =========================================================
    const customFromRaw = fromDate;
    const customToRaw = toDate;

    let customFrom = toDbDateOnly(customFromRaw);
    let customTo = toDbDateOnly(customToRaw);

    const useCustom = !!(customFrom && customTo);

    if (useCustom && customFrom > customTo) {
      const tmp = customFrom;
      customFrom = customTo;
      customTo = tmp;
    }

    let beforeRows = [];
    let afterRows = [];

    if (useCustom) {
      const allRows = await InsuRenewal.findAll({
        where: {
          ...baseWhere,
          POLICY_END_DATE: { [Op.between]: [customFrom, customTo] },
        },
        raw: true,
      });

      beforeRows = allRows
        .map((r) => ({ ...r, __END: ensureYmdOnly(r.POLICY_END_DATE) }))
        .filter((r) => r.__END && r.__END >= today)
        .sort((a, b) => String(a.__END).localeCompare(String(b.__END)));

      afterRows = allRows
        .map((r) => ({ ...r, __END: ensureYmdOnly(r.POLICY_END_DATE) }))
        .filter((r) => r.__END && r.__END < today)
        .sort((a, b) => String(b.__END).localeCompare(String(a.__END)));
    } else {
      beforeRows = await InsuRenewal.findAll({
        where: {
          ...baseWhere,
          POLICY_END_DATE: { [Op.between]: [today, beforeTo] },
        },
        order: [["POLICY_END_DATE", "ASC"]],
        raw: true,
      });

      afterRows = await InsuRenewal.findAll({
        where: {
          ...baseWhere,
          POLICY_END_DATE: { [Op.between]: [afterFrom, yesterday] },
        },
        order: [["POLICY_END_DATE", "DESC"]],
        raw: true,
      });
    }

    // Attach latest followup details to rows so status can be shown/used in frontend
    const attachFollowup = async (rows) => {
      return Promise.all(
        rows.map(async (r) => {
          const fu = await sequelize.query(
            `SELECT TOP 1 UTD, FOLLOWUP_STATUS, 
                    CONVERT(varchar(10), FOLLOWUP_DATE, 120) AS FOLLOWUP_DATE, 
                    FOLLOWUP_TIME, REMARKS, CREATED_AT
             FROM dbo.FOLLOWUP_DETAILS
             WHERE TRAN_ID = :tranId
             ORDER BY UTD DESC`,
            {
              replacements: { tranId: r.TRAN_ID ?? null },
              type: sequelize.QueryTypes.SELECT,
            }
          );
          return {
            ...r,
            FOLLOWUP_UTD: fu?.[0]?.UTD || null,
            FOLLOWUP_STATUS: fu?.[0]?.FOLLOWUP_STATUS || null,
            FOLLOWUP_DATE: fu?.[0]?.FOLLOWUP_DATE || null,
            FOLLOWUP_TIME: fu?.[0]?.FOLLOWUP_TIME || null,
            REMARKS: fu?.[0]?.REMARKS || null,
            FOLLOWUP_CREATED_AT: fu?.[0]?.CREATED_AT || null,
          };
        })
      );
    };

    beforeRows = await attachFollowup(beforeRows);
    afterRows = await attachFollowup(afterRows);

    const beforeReminders = beforeRows.map((r) => {
      const endYmd = ensureYmdOnly(r.POLICY_END_DATE) || today;
      const daysLeft = diffDays(today, endYmd);

      const showMsg = !useCustom || (daysLeft >= 0 && daysLeft <= beforeDays);

      return {
        type: "BEFORE_EXPIRY",
        days: daysLeft,
        message: showMsg
          ? `Your policy is going to expire in ${daysLeft} day(s). Please renew it.`
          : "",
        data: r,
      };
    });

    const afterReminders = afterRows.map((r) => {
      const endYmd = ensureYmdOnly(r.POLICY_END_DATE) || today;
      const daysAgo = diffDays(endYmd, today);

      const showMsg = !useCustom || (daysAgo > 0 && daysAgo <= afterDays);

      return {
        type: "AFTER_EXPIRY",
        days: daysAgo,
        message: showMsg
          ? `Your policy has expired ${daysAgo} day(s) ago. Please renew it.`
          : "",
        data: r,
      };
    });

    return res.json({
      ok: true,
      config: {
        beforeDays,
        afterDays,
        exportType: exportTypeFinal,
        loc_code: locCodes,
        isAdmin: emp_dms_code && String(emp_dms_code).trim() !== "",
      },
      range: useCustom
        ? { customFrom, customTo, today }
        : {
          today,
          beforeReminderTo: beforeTo,
          afterReminderFrom: afterFrom,
          afterReminderTo: yesterday,
        },
      counts: { before: beforeReminders.length, after: afterReminders.length },
      beforeReminders,
      afterReminders,
      debug: {
        blockedVehCount: blockedVehKeys.length,
        branchFiltered: !!locCodes.length,
        employeeFiltered: employeePlaceholders !== "",
      },
    });
  } catch (err) {
    console.error("Error in getInsuRenewalReminders:", err);
    return res.status(500).json({
      ok: false,
      Message: err.message || "Something went wrong",
    });
  } finally {
    await sequelize.close();
  }
};


exports.saveInsuRenewalFollowupByVehicle = async function (req, res, next) {
  const sequelize = await dbname(req, req.headers.compcode);
  const t = await sequelize.transaction();

  try {
    const InsuRenewalMst = Insu_Renewal_Mst(sequelize, DataTypes);
    const FollowupDetail = FollowupDetailModel(sequelize, DataTypes);

    const pad2 = (x) => String(x).padStart(2, "0");
    const ymd = (d = new Date()) => {
      const yyyy = d.getFullYear();
      const mm = pad2(d.getMonth() + 1);
      const dd = pad2(d.getDate());
      return `${yyyy}-${mm}-${dd}`;
    };

    // ✅ Now supports:
    //  - DD/MM/YYYY
    //  - DD-MM-YYYY
    //  - YYYY-MM-DD
    //  - YYYY-MM-DDTHH:mm:ss...
    const toDbDateOnly = (v) => {
      if (!v) return null;
      const s0 = String(v).trim();
      if (!s0) return null;

      // ISO date/time -> take first 10
      if (/^\d{4}-\d{2}-\d{2}/.test(s0)) return s0.slice(0, 10);

      let m = s0.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s0.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s0.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;

      return null;
    };

    const normalizeVehNo = (v) => {
      if (!v) return "";
      return String(v).toUpperCase().replace(/\s+/g, "").trim();
    };

    // ✅ FOLLOWUP_TIME parser:
    // supports:
    //  - HH:mm
    //  - HH:mm:ss
    //  - 7:30 PM / 07:30 PM
    // returns "HH:mm:ss" (SQL time compatible)
    const toDbTimeOnly = (v) => {
      if (v === undefined || v === null) return null;
      const s0 = String(v).trim();
      if (!s0) return null;

      // HH:mm or HH:mm:ss (24h)
      let m = s0.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (m) {
        const hh = Number(m[1]);
        const mm = Number(m[2]);
        const ss = Number(m[3] || 0);
        if (hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59 && ss >= 0 && ss <= 59) {
          return `${pad2(hh)}:${pad2(mm)}:${pad2(ss)}`;
        }
        return null;
      }

      // h:mm AM/PM
      m = s0.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
      if (m) {
        let hh = Number(m[1]);
        const mm = Number(m[2]);
        const ap = String(m[3]).toUpperCase();
        if (hh < 1 || hh > 12 || mm < 0 || mm > 59) return null;
        if (ap === "PM" && hh !== 12) hh += 12;
        if (ap === "AM" && hh === 12) hh = 0;
        return `${pad2(hh)}:${pad2(mm)}:00`;
      }

      return null;
    };

    // ✅ MST EXPORT_TYPE: 1 latest, 33 old
    const exportType = parseInt(req.body?.EXPORT_TYPE ?? 1, 10);
    const exportTypeFinal = Number.isFinite(exportType) ? exportType : 1;

    const vehRaw =
      req.body?.VEHICAL_REG_NO ||
      req.body?.vehical_reg_no ||
      req.body?.vehicleNo;
    const vehKey = normalizeVehNo(vehRaw);

    const FOLLOWUP_DATE = toDbDateOnly(
      req.body?.FOLLOWUP_DATE || req.body?.followupDate,
    );

    // ✅ NEW: FOLLOWUP_TIME
    const FOLLOWUP_TIME = toDbTimeOnly(
      req.body?.FOLLOWUP_TIME || req.body?.followupTime,
    );

    const FOLLOWUP_STATUS = String(req.body?.FOLLOWUP_STATUS || "PROMISED")
      .trim()
      .toUpperCase()
      .slice(0, 100);

    const REMARKS = String(req.body?.REMARKS || req.body?.remarks || "")
      .trim()
      .slice(0, 100);

    if (!vehKey) {
      await t.rollback();
      return res
        .status(400)
        .json({ ok: false, Message: "VEHICAL_REG_NO is required" });
    }
    if (!FOLLOWUP_DATE) {
      await t.rollback();
      return res
        .status(400)
        .json({ ok: false, Message: "FOLLOWUP_DATE is required" });
    }

    // ✅ optional validation: if user sent time but parsing failed
    if ((req.body?.FOLLOWUP_TIME || req.body?.followupTime) && !FOLLOWUP_TIME) {
      await t.rollback();
      return res.status(400).json({
        ok: false,
        Message:
          "Invalid FOLLOWUP_TIME. Supported: HH:mm, HH:mm:ss, h:mm AM/PM",
      });
    }

    // ✅ 1) Find MST by vehicle + export type (locking to prevent duplicates)
    const mstRows = await sequelize.query(
      `
      SELECT TOP 1
        UTD,
        VEHICAL_REG_NO,
        EXPORT_TYPE
      FROM dbo.INSU_RENEWAL_MST WITH (UPDLOCK, HOLDLOCK)
      WHERE REPLACE(UPPER(VEHICAL_REG_NO), ' ', '') = :vehKey
        AND EXPORT_TYPE = :exportType
      ORDER BY UTD DESC
      `,
      {
        type: sequelize.QueryTypes.SELECT,
        replacements: { vehKey, exportType: exportTypeFinal },
        transaction: t,
      },
    );

    let mst = mstRows?.[0] || null;

    // ✅ 2) If not found -> create MST entry
    if (!mst) {
      const createMstPayload = {
        VEHICAL_REG_NO: vehRaw,
        EXPORT_TYPE: exportTypeFinal,
        CREATED_AT: new Date(),
        VALIDFROM: new Date(),
        VALIDTO: "9999-12-31",
      };

      const createdMst = await InsuRenewalMst.create(createMstPayload, {
        transaction: t,
      });
      mst = createdMst?.get ? createdMst.get({ plain: true }) : createdMst;

      if (!mst?.UTD) {
        await t.rollback();
        return res.status(500).json({
          ok: false,
          Message: "Failed to create master entry (UTD not generated)",
        });
      }
    }

    const utd = parseInt(mst.UTD, 10);
    if (!Number.isFinite(utd)) {
      await t.rollback();
      return res
        .status(500)
        .json({ ok: false, Message: "Invalid UTD in master" });
    }

    // ✅ 3) Find latest followup for this TRAN_ID + same export type
    const current = await FollowupDetail.findOne({
      where: { TRAN_ID: utd, EXPORT_TYPE: exportTypeFinal },
      order: [["FOLLOWUP_DATE", "DESC"]],
      raw: true,
      transaction: t,
    });

    // ✅ Optional: prevent exact duplicates (same date + time + status + remarks)
    if (
      current &&
      String(current.FOLLOWUP_DATE || "").slice(0, 10) === FOLLOWUP_DATE &&
      String(current.FOLLOWUP_TIME || "").slice(0, 8) ===
      String(FOLLOWUP_TIME || "").slice(0, 8) &&
      String(current.FOLLOWUP_STATUS || "").toUpperCase() === FOLLOWUP_STATUS &&
      String(current.REMARKS || "") === REMARKS
    ) {
      await t.commit();
      return res.json({
        ok: true,
        Message: "Same follow-up already exists (no new entry created)",
        link: {
          TRAN_ID: utd,
          UTD: utd,
          VEHICAL_REG_NO: mst.VEHICAL_REG_NO,
          MST_EXPORT_TYPE: exportTypeFinal,
        },
        mst,
        data: current,
      });
    }

    // ✅ 4) Insert new followup row (history kept)
    const newRow = await FollowupDetail.create(
      {
        TRAN_ID: utd,
        EXPORT_TYPE: exportTypeFinal,
        FOLLOWUP_STATUS,
        FOLLOWUP_DATE,
        FOLLOWUP_TIME, // ✅ ADDED
        LAST_FOLLOWUP_DATE: current?.FOLLOWUP_DATE
          ? String(current.FOLLOWUP_DATE).slice(0, 10)
          : FOLLOWUP_DATE,
        REMARKS,
      },
      { transaction: t },
    );

    await t.commit();
    return res.json({
      ok: true,
      Message: current
        ? "Follow-up saved (new entry created)"
        : "Follow-up created (first entry)",
      link: {
        TRAN_ID: utd,
        UTD: utd,
        VEHICAL_REG_NO: mst.VEHICAL_REG_NO,
        MST_EXPORT_TYPE: exportTypeFinal,
      },
      mst,
      oldFollowup: current || null,
      data: newRow,
    });
  } catch (err) {
    try {
      await t.rollback();
    } catch (_) { }
    return res.status(500).json({ ok: false, Message: err.message });
  } finally {
    await sequelize.close();
  }
};


exports.getFollowupHistoryByVehicle = async function (req, res, next) {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const normalizeVehNo = (v) =>
      String(v || "")
        .toUpperCase()
        .replace(/\s+/g, "")
        .trim();

    const vehRaw =
      req.body?.VEHICAL_REG_NO ||
      req.body?.vehical_reg_no ||
      req.body?.vehicleNo;

    const vehKey = normalizeVehNo(vehRaw);

    if (!vehKey) {
      return res
        .status(400)
        .json({ ok: false, Message: "VEHICAL_REG_NO is required" });
    }

    // optional: if you want history according to master export type (1 latest / 33 old)
    const exportType = parseInt(req.body?.EXPORT_TYPE ?? 1, 10);
    const exportTypeFinal = Number.isFinite(exportType) ? exportType : 1;

    // ✅ NOTE:
    // Tumhare reminders controller me table name "FOLLOWUP_DETAILS" use ho raha hai,
    // isliye yaha bhi same use kiya hai.
    // Agar tumhara actual table "FOLLOWUP_DETAIL" hai to dono jagah same naam kar do.

    const rows = await sequelize.query(
      `
        SELECT
          f.TRAN_ID,
          m.UTD AS MST_UTD,
          m.VEHICAL_REG_NO,
          CONVERT(varchar(10), f.FOLLOWUP_DATE, 120) AS FOLLOWUP_DATE,
          CONVERT(varchar(10), f.LAST_FOLLOWUP_DATE, 120) AS LAST_FOLLOWUP_DATE,
          f.FOLLOWUP_STATUS,
          f.REMARKS
        FROM dbo.FOLLOWUP_DETAILS f
        INNER JOIN dbo.INSU_RENEWAL_MST m
          ON m.UTD = f.TRAN_ID
        WHERE REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '') = :vehKey
          AND m.EXPORT_TYPE = :exportType
        ORDER BY
          f.FOLLOWUP_DATE DESC
        `,
      {
        type: sequelize.QueryTypes.SELECT,
        replacements: { vehKey, exportType: exportTypeFinal },
      },
    );

    return res.json({
      ok: true,
      count: rows.length,
      data: rows,
    });
  } catch (err) {
    return res.status(500).json({ ok: false, Message: err.message });
  } finally {
    await sequelize.close();
  }
};

exports.getVehicleByRegNo = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);

  // ✅ FIX: loc_code body ya query se lo, aur undefined ho to WHERE me mat bhejo
  const loc_code = req.body?.loc_code ?? req.query?.loc_code;

  console.log("loc_code", loc_code);

  try {
    // ✅ MODELS
    const InsuRenewal = Insu_Renewal(sequelize, DataTypes); // dbo.INSU_RENEWAL (detail)
    const InsuRenewalMst = Insu_Renewal_Mst(sequelize, DataTypes); // dbo.INSU_RENEWAL_MST (master)
    const InsuRenewalPymt = Insu_Renewal_Pymt(sequelize, DataTypes); // dbo.INSU_RENEWAL_PYMT (payment)

    const regNoRaw = req.body?.RegNo || req.query?.RegNo; // POST body OR GET query
    const regNo = String(regNoRaw || "").trim().toUpperCase();

    if (!regNo) {
      return res.status(400).json({ Status: false, Message: "RegNo required" });
    }

    // ✅ FIX: undefined loc_code ko where me pass mat karo (warna Sequelize error deta hai)
    const mstWhere = { VEHICAL_REG_NO: regNo };
    if (loc_code !== undefined && loc_code !== null && String(loc_code).trim() !== "") {
      mstWhere.loc_code = loc_code;
    }

    // 1) MASTER (vehicle -> master)
    const mst = await InsuRenewalMst.findOne({
      where: mstWhere,
      order: [["UTD", "DESC"]],
      raw: true,
    });

    if (!mst) {
      return res.json({ Status: true, Result: [], Message: "New vehicle" });
    }

    // ✅ LINK KEY
    // INSU_RENEWAL.TRAN_ID      === INSU_RENEWAL_MST.UTD
    // INSU_RENEWAL_PYMT.TRAN_ID === INSU_RENEWAL_MST.UTD
    const tranId = mst.UTD;

    // 2) INSURANCE DETAIL (latest) from INSU_RENEWAL by TRAN_ID
    // Prefer EXPORT_TYPE=1 (active/latest)
    let insuDetail = await InsuRenewal.findOne({
      where: { TRAN_ID: tranId, EXPORT_TYPE: 1 },
      order: [["UTD", "DESC"]],
      raw: true,
    });

    // fallback: system_time all (latest by VALIDFROM) if needed
    if (!insuDetail) {
      const [rows] = await sequelize.query(
        `
        SELECT TOP 1 *
        FROM dbo.INSU_RENEWAL FOR SYSTEM_TIME ALL
        WHERE TRAN_ID = :tranId
        ORDER BY VALIDFROM DESC
        `,
        { replacements: { tranId } }
      );
      insuDetail = rows?.[0] || null;
    }

    // 3) PAYMENT DETAIL (latest) from INSU_RENEWAL_PYMT by TRAN_ID
    let pymtDetail = await InsuRenewalPymt.findOne({
      where: { TRAN_ID: tranId },
      order: [
        ["PYMT_DATE", "DESC"],
        ["UTD", "DESC"],
      ],
      raw: true,
    });

    // fallback: system_time all
    if (!pymtDetail) {
      const [rows] = await sequelize.query(
        `
        SELECT TOP 1 *
        FROM dbo.INSU_RENEWAL_PYMT FOR SYSTEM_TIME ALL
        WHERE TRAN_ID = :tranId
        ORDER BY VALIDFROM DESC
        `,
        { replacements: { tranId } }
      );
      pymtDetail = rows?.[0] || null;
    }

    // ✅ Frontend-friendly mapping (same structure, now populated from INSU_RENEWAL)
    const resultObj = {
      VehicleId: mst?.UTD ?? null,
      RegNo: regNo,

      OwnerName: insuDetail?.CUST_NAME ?? "",
      MobileNo: insuDetail?.CUST_MOB_NO ? String(insuDetail.CUST_MOB_NO) : "",
      ModelVariant: insuDetail?.MODEL_NAME ?? "",

      PolicyNo: insuDetail?.POLICY_NUMBER ? String(insuDetail.POLICY_NUMBER) : "",
      PolicyName: insuDetail?.POLICY_NAME ?? "",
      PolicyStartDate: insuDetail?.POLICY_START_DATE ?? "",
      PolicyExpiryDate: insuDetail?.POLICY_END_DATE ?? "",

      PrevPolicyNo: insuDetail?.POLICY_NUMBER ? String(insuDetail.POLICY_NUMBER) : "",
      PrevExpiryDate: insuDetail?.POLICY_END_DATE ?? "",
      PrevInsuranceCo: "",

      Payment: pymtDetail
        ? {
          UTD: pymtDetail.UTD ?? null,
          TRAN_ID: pymtDetail.TRAN_ID ?? tranId,

          INSU_TYPE: pymtDetail.INSU_TYPE ?? null,
          PREMIUM_AMOUNT: pymtDetail.PREMIUM_AMOUNT ?? null,

          PYMT_MODE: pymtDetail.PYMT_MODE ?? "",
          PYMT_DATE: pymtDetail.PYMT_DATE ?? "",
          PYMT_AMOUNT: pymtDetail.PYMT_AMOUNT ?? null,

          // aapke table me column PYMT_EMPCODE hai
          PYMT_EMPCODE: pymtDetail.PYMT_EMPCODE ?? "",
          PYMT_REMARK: pymtDetail.PYMT_REMARK ?? "",
          PYMT_STATUS: pymtDetail.PYMT_STATUS ?? null,

          ACNT_APPR_CODE: pymtDetail.ACNT_APPR_CODE ?? "",
          ACNT_APPR_REMARK: pymtDetail.ACNT_APPR_REMARK ?? "",
          ACNT_APPR_STATUS: pymtDetail.ACNT_APPR_STATUS ?? null,
          ACNT_APPR_DATE: pymtDetail.ACNT_APPR_DATE ?? "",

          POLICY_GENERATE: pymtDetail.POLICY_GENERATE ?? null,
          UTR: pymtDetail.UTR ?? null,
          BANK_NAME: pymtDetail.BANK_NAME ?? "",
          REMARKS: pymtDetail.REMARKS ?? "",

          DOC_PATH: pymtDetail.DOC_PATH ?? "",
          CREATED_AT: pymtDetail.CREATED_AT ?? "",
        }
        : null,
    };

    return res.json({
      Status: true,
      Result: [resultObj],
      Message: "Vehicle found",
    });
  } catch (e) {
    return res.status(500).json({ Status: false, Message: e.message });
  } finally {
    await sequelize.close();
  }
};

exports.getInsuranceAndPaymentDropdowns = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    // ✅ Only active rows (Export_Type=33 excluded)
    const rows = await sequelize.query(
      `
      SELECT
        UTD,
        Misc_Type,
        Misc_Code,
        Misc_Name,
        Misc_Abbr,
        Misc_Add1, Misc_Add2, Misc_Add3,
        Misc_Dtl1, Misc_Dtl2, Misc_Dtl3,
        Misc_Num1, MISC_NUM2,
        Export_Type
      FROM dbo.MISC_MST
      WHERE Misc_Type IN (9, 18)
        AND ISNULL(Export_Type, 1) <> 33
      ORDER BY Misc_Type, Misc_Code
      `,
      { type: Sequelize.QueryTypes.SELECT }
    );

    const toOption = (r) => ({
      label: String(r.Misc_Name ?? ""),
      value: String(r.Misc_Abbr || r.Misc_Name || r.Misc_Code || r.UTD),
      code: r.Misc_Code,
      utd: r.UTD,
      meta: r,
    });

    const insuranceCompanies = rows
      .filter((r) => Number(r.Misc_Type) === 9)
      .map(toOption);

    const paymentModes = rows
      .filter((r) => Number(r.Misc_Type) === 18)
      .map(toOption);

    return res.json({
      Status: true,
      Message: "Dropdowns loaded",
      Result: { insuranceCompanies, paymentModes },
    });
  } catch (e) {
    return res.status(500).json({ Status: false, Message: e.message });
  } finally {
    await sequelize.close();
  }
};

exports.SaveInsuranceRenewal = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const InsuRenewal = Insu_Renewal(sequelize, DataTypes);
    const InsuRenewalMst = Insu_Renewal_Mst(sequelize, DataTypes);
    const InsuRenewalPymt = Insu_Renewal_Pymt(sequelize, DataTypes);

    const b = req.body || {};

    const regNo = String(b.RegNo || "").trim().toUpperCase();
    if (!regNo) return res.status(400).json({ Status: false, Message: "RegNo required" });

    // ✅ FIX: loc_code resolve (body -> query -> headers -> user) and parse integer
    const rawLoc =
      b?.loc_code ??
      b?.Loc_Code ??
      b?.LOC_CODE ??
      b?.branch ??
      req.body?.loc_code ??
      req.body?.Loc_Code ??
      req.body?.LOC_CODE ??
      req.body?.branch ??
      req.query?.loc_code ??
      req.headers?.loc_code ??
      req.headers?.branch ??
      req.user?.branch ??
      null;

    const loc_code =
      rawLoc != null && String(rawLoc).trim() !== "" && !isNaN(rawLoc)
        ? parseInt(rawLoc, 10)
        : null;
    console.log("loc_code", loc_code);

    // =========================================================
    // ✅ FIX: Resolve logged-in EMPCODE reliably
    // Priority:
    // 1) req.user.EMPCODE (if middleware attaches decoded token with EMPCODE)
    // 2) headers/body empcode
    // 3) fetch from dbo.user_tbl using req.user.User_Code / User_Name / headers.name
    // =========================================================
    const resolveEmpCode = async () => {
      let emp =
        req.user?.Emp_Code ??
        req.user?.EMPCODE ??
        req.user?.empcode ??
        req.headers?.empcode ??
        req.headers?.emp_code ??
        b.EmpCode ??
        b.empcode ??
        null;

      emp = emp != null ? String(emp).trim() : "";
      if (emp && emp !== "0") return emp;

      const userCode = req.user?.User_Code ?? req.user?.UserCode ?? null;
      const userName =
        req.user?.User_Name ??
        req.user?.UserName ??
        req.headers?.name ??
        req.headers?.user_name ??
        null;

      // If we have userCode or userName, lookup from user_tbl
      if (!userCode && !userName) return null;

      const rows = await sequelize.query(
        `
        SELECT TOP 1 EMPCODE
        FROM dbo.user_tbl
        WHERE Export_Type = 1
          AND Module_Code = 10
          AND (
            (:userCode IS NOT NULL AND User_Code = :userCode)
            OR
            (:userName IS NOT NULL AND UPPER(LTRIM(RTRIM(User_Name))) = UPPER(LTRIM(RTRIM(:userName))))
          )
        ORDER BY User_Code DESC
        `,
        {
          type: sequelize.QueryTypes.SELECT,
          replacements: {
            userCode: userCode != null ? Number(userCode) : null,
            userName: userName != null ? String(userName) : null,
          },
        }
      );

      const dbEmp = rows?.[0]?.EMPCODE != null ? String(rows[0].EMPCODE).trim() : "";
      if (dbEmp && dbEmp !== "0") return dbEmp;

      return null;
    };

    const empCode = await resolveEmpCode();
    if (!empCode) {
      return res.status(401).json({
        Status: false,
        Message: "EMPCODE not found for logged-in user. Please login again (or ensure auth middleware sets req.user).",
      });
    }

    // multer file
    const file = req.file;
    if (!file) return res.status(400).json({ Status: false, Message: "PaymentProof file required" });

    const rawComp = req.headers.compcode || req.body.compcode || "";
    const compCode = String(rawComp).trim().split("-")[0];
    if (!compCode) return res.status(400).json({ Status: false, Message: "compcode header required" });

    // ✅ Upload proof
    const uploaded = await uploadPaymentProofSMB(file, compCode);
    const docUrl = uploaded.url;

    const onlyDigits = (v) => {
      const s = String(v ?? "").replace(/\D/g, "");
      return s ? s : null;
    };

    const toDateOnly = (v) => {
      if (!v) return null;
      const s = String(v).trim();
      if (!s) return null;

      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

      let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return `${m[3]}-${String(m[2]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;

      m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
      if (m) return `${m[3]}-${String(m[2]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;

      return null;
    };

    const resolveInsuranceName = async (input, transaction) => {
      if (input == null) return null;
      const s = String(typeof input === "object" ? (input.value || input.label || "") : input).trim();
      if (!s) return null;

      const rows = await sequelize.query(
        `
        SELECT TOP 1 Misc_Name, Misc_Abbr, Misc_Code, UTD
        FROM dbo.MISC_MST
        WHERE Misc_Type = 9
          AND ISNULL(Export_Type, 1) <> 33
          AND (
            (:utdNum IS NOT NULL AND (UTD = :utdNum OR Misc_Code = :utdNum))
            OR UPPER(LTRIM(RTRIM(Misc_Abbr))) = UPPER(:k)
            OR UPPER(LTRIM(RTRIM(Misc_Name))) = UPPER(:k)
          )
        ORDER BY UTD DESC
        `,
        {
          type: sequelize.QueryTypes.SELECT,
          replacements: {
            k: s,
            utdNum: /^\d+$/.test(s) ? parseInt(s, 10) : null,
          },
          transaction,
        }
      );

      if (rows && rows.length > 0 && rows[0].Misc_Name) {
        return String(rows[0].Misc_Name).trim();
      }
      return s;
    };

    const resolvePaymentModeName = async (input, transaction) => {
      if (input == null) return null;
      const s = String(typeof input === "object" ? (input.value || input.label || "") : input).trim();
      if (!s) return null;

      const rows = await sequelize.query(
        `
        SELECT TOP 1 Misc_Name, Misc_Abbr, Misc_Code, UTD
        FROM dbo.MISC_MST
        WHERE Misc_Type = 18
          AND ISNULL(Export_Type, 1) <> 33
          AND (
            (:utdNum IS NOT NULL AND (UTD = :utdNum OR Misc_Code = :utdNum))
            OR UPPER(LTRIM(RTRIM(Misc_Abbr))) = UPPER(:k)
            OR UPPER(LTRIM(RTRIM(Misc_Name))) = UPPER(:k)
          )
        ORDER BY UTD DESC
        `,
        {
          type: sequelize.QueryTypes.SELECT,
          replacements: {
            k: s,
            utdNum: /^\d+$/.test(s) ? parseInt(s, 10) : null,
          },
          transaction,
        }
      );

      if (rows && rows.length > 0) {
        return String(rows[0].Misc_Abbr || rows[0].Misc_Name).trim();
      }
      return s;
    };

    // ✅ Resolve insurance company name and payment mode
    const insuType = await resolveInsuranceName(b.InsuranceCo ?? b.INSU_TYPE, null);
    if (!insuType) {
      return res.status(400).json({ Status: false, Message: "InsuranceCo required." });
    }

    const pymtModeType = await resolvePaymentModeName(b.PaymentMode ?? b.PYMT_MODE, null);
    if (!pymtModeType) {
      return res.status(400).json({ Status: false, Message: "PaymentMode required." });
    }

    const pymtDate = toDateOnly(b.PaymentDate ?? b.PYMT_DATE);
    if (!pymtDate) return res.status(400).json({ Status: false, Message: "PaymentDate required" });

    // ---- transaction ----
    let createdInsu = null;
    let createdPymt = null;

    const cleanReg = regNo.replace(/[\s\-\/]+/g, "").toUpperCase();

    await sequelize.transaction(async (t) => {
      // 1) Find existing master by VEHICAL_REG_NO (regardless of loc_code differences)
      const mstRows = await sequelize.query(
        `SELECT TOP 1 UTD, VEHICAL_REG_NO, LOC_CODE 
         FROM dbo.INSU_RENEWAL_MST 
         WHERE REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(VEHICAL_REG_NO, '')))),' ',''),'-',''),'/','') = :cleanReg
         ORDER BY UTD ASC`,
        {
          replacements: { cleanReg },
          type: sequelize.QueryTypes.SELECT,
          transaction: t,
        }
      );

      let tranId = mstRows?.[0]?.UTD || null;

      if (!tranId) {
        const newMst = await InsuRenewalMst.create(
          {
            VEHICAL_REG_NO: regNo,
            LOC_CODE: loc_code != null ? loc_code : null,
            EXPORT_TYPE: 1,
            CREATED_AT: Sequelize.literal("GETDATE()"),
          },
          {
            transaction: t,
            validate: false,
            fields: ["VEHICAL_REG_NO", "LOC_CODE", "EXPORT_TYPE", "CREATED_AT"],
          }
        );
        console.log("newMst", newMst);
        tranId = newMst?.UTD ?? newMst?.dataValues?.UTD ?? null;

        // ✅ Direct SQL guarantee to save LOC_CODE in DB
        if (tranId && loc_code != null) {
          await sequelize.query(
            `UPDATE dbo.INSU_RENEWAL_MST
             SET LOC_CODE = :locCode, EXPORT_TYPE = 1
             WHERE UTD = :utd`,
            {
              replacements: { locCode: loc_code, utd: tranId },
              transaction: t,
            }
          );
        }
      } else {
        const existingLocCode = mstRows[0]?.LOC_CODE ?? mstRows[0]?.loc_code ?? null;
        if (loc_code != null && (existingLocCode == null || String(existingLocCode).trim() === "")) {
          await InsuRenewalMst.update(
            { LOC_CODE: loc_code },
            { where: { UTD: tranId }, transaction: t }
          );
          await sequelize.query(
            `UPDATE dbo.INSU_RENEWAL_MST
             SET LOC_CODE = :locCode, EXPORT_TYPE = 1
             WHERE UTD = :utd`,
            {
              replacements: { locCode: loc_code, utd: tranId },
              transaction: t,
            }
          );
        }
      }

      if (!tranId) throw new Error("Failed to get TRAN_ID from master");

      // 2) Find existing active INSU_RENEWAL row for this vehicle
      const existingInsuRows = await sequelize.query(
        `SELECT TOP 1 UTD 
         FROM dbo.INSU_RENEWAL 
         WHERE (TRAN_ID = :tranId OR REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(VEHICAL_REG_NO, '')))),' ',''),'-',''),'/','') = :cleanReg)
           AND ISNULL(EXPORT_TYPE, 1) = 1
         ORDER BY UTD DESC`,
        {
          replacements: { tranId, cleanReg },
          type: sequelize.QueryTypes.SELECT,
          transaction: t,
        }
      );

      const insuObj = {
        VEHICAL_REG_NO: regNo,
        CUST_NAME: String(b.OwnerName || "").trim(),
        CUST_MOB_NO: onlyDigits(b.MobileNo),
        MODEL_NAME: String(b.ModelVariant || "").trim() || null,

        POLICY_NAME: String(b.PolicyName || "").trim(),
        POLICY_NUMBER: b.PolicyNo ? String(b.PolicyNo).trim() : null,

        POLICY_START_DATE: b.PolicyStartDate || null,
        POLICY_END_DATE: b.PolicyExpiryDate || null,

        INSU_TYPE: insuType,
        PREMIUM_AMOUNT: b.PremiumAmount || null,

        PAYMENT_MODE: String(pymtModeType),
        PAYMENT_DATE: pymtDate,
        PAYMENT_AMOUNT: b.Amount || null,
        UTR: b.UTR || null,
        CHEQUE_NO: b.ChequeNo || null,
        BANK_NAME: b.BankName || null,
        REMARKS: b.Remarks || null,

        DOC_PATH: docUrl,
        TRAN_ID: tranId,
        EXPORT_TYPE: 1,
        CREATED_AT: Sequelize.literal("GETDATE()"),
      };

      if (existingInsuRows && existingInsuRows.length > 0) {
        const activeInsuUtd = existingInsuRows[0].UTD;
        await InsuRenewal.update(insuObj, {
          where: { UTD: activeInsuUtd },
          transaction: t,
        });
        createdInsu = { UTD: activeInsuUtd };
      } else {
        createdInsu = await InsuRenewal.create(insuObj, {
          transaction: t,
          validate: false,
          fields: [
            "VEHICAL_REG_NO",
            "CUST_NAME",
            "CUST_MOB_NO",
            "MODEL_NAME",
            "POLICY_NAME",
            "POLICY_NUMBER",
            "POLICY_START_DATE",
            "POLICY_END_DATE",
            "INSU_TYPE",
            "PREMIUM_AMOUNT",
            "PAYMENT_MODE",
            "PAYMENT_DATE",
            "PAYMENT_AMOUNT",
            "UTR",
            "CHEQUE_NO",
            "BANK_NAME",
            "REMARKS",
            "DOC_PATH",
            "TRAN_ID",
            "EXPORT_TYPE",
            "CREATED_AT",
          ],
        });
      }

      // 3) Find existing payment record in INSU_RENEWAL_PYMT for this vehicle / TRAN_ID
      const existingPymtRows = await sequelize.query(
        `SELECT p.UTD, p.TRAN_ID 
         FROM dbo.INSU_RENEWAL_PYMT p
         LEFT JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = p.TRAN_ID
         WHERE p.TRAN_ID = :tranId
            OR (m.VEHICAL_REG_NO IS NOT NULL AND REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(m.VEHICAL_REG_NO))),' ',''),'-',''),'/','') = :cleanReg)
         ORDER BY p.UTD ASC`,
        {
          replacements: { tranId, cleanReg },
          type: sequelize.QueryTypes.SELECT,
          transaction: t,
        }
      );

      const pymtDataObj = {
        TRAN_ID: tranId,
        INSU_TYPE: insuType,
        PREMIUM_AMOUNT: b.PremiumAmount || null,

        PYMT_MODE: String(pymtModeType),
        PYMT_DATE: pymtDate,
        PYMT_AMOUNT: b.Amount || null,

        PYMT_CODE: empCode,

        PYMT_REMARK: b.Remarks || null,
        PYMT_STATUS: 1,

        UTR: b.UTR || null,
        BANK_NAME: b.BankName || null,
        REMARKS: b.Remarks || null,

        DOC_PATH: docUrl,
        CREATED_AT: Sequelize.literal("GETDATE()"),
      };

      if (existingPymtRows && existingPymtRows.length > 0) {
        const primaryUtd = existingPymtRows[0].UTD;
        await InsuRenewalPymt.update(
          {
            ...pymtDataObj,
            ACNT_APPR_STATUS: null,
            ACNT_APPR_CODE: null,
            ACNT_APPR_REMARK: null,
            ACNT_APPR_DATE: null,
          },
          {
            where: { UTD: primaryUtd },
            transaction: t,
          }
        );
        createdPymt = { UTD: primaryUtd };

        // Clean up duplicate historical payment rows for this vehicle if any
        if (existingPymtRows.length > 1) {
          const dupUtds = existingPymtRows.slice(1).map((r) => r.UTD);
          await InsuRenewalPymt.destroy({
            where: { UTD: dupUtds },
            transaction: t,
          });
        }
      } else {
        createdPymt = await InsuRenewalPymt.create(pymtDataObj, {
          transaction: t,
          validate: false,
          fields: [
            "TRAN_ID",
            "INSU_TYPE",
            "PREMIUM_AMOUNT",
            "PYMT_MODE",
            "PYMT_DATE",
            "PYMT_AMOUNT",
            "PYMT_CODE",
            "PYMT_REMARK",
            "PYMT_STATUS",
            "UTR",
            "BANK_NAME",
            "REMARKS",
            "DOC_PATH",
            "CREATED_AT",
          ],
        });
      }
    });

    return res.json({
      Status: true,
      Message: "Insurance renewal and payment details have been recorded successfully.",
      Result: {
        VEHICAL_REG_NO: regNo,
        INSU_UTD: createdInsu?.UTD ?? createdInsu?.dataValues?.UTD ?? null,
        PYMT_UTD: createdPymt?.UTD ?? createdPymt?.dataValues?.UTD ?? null,
        PYMT_CODE: empCode,
      },
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ Status: false, Message: e.message });
  } finally {
    await sequelize.close();
  }
};


const FILE_UPLOAD_BASE_URL = "https://erp.autovyn.com/backend";
const FILE_FETCH_BASE_URL = "https://erp.autovyn.com/backend/fetch?filePath=";

async function uploadPaymentProofSMB(file, compCode) {
  const customPath = `${compCode}/insurance/`;

  const original = String(file.originalname || "proof.png").replace(
    /\s+/g,
    "_",
  );
  const filename = `${Date.now()}_${original}`;

  const formData = new FormData();
  formData.append("photo", file.buffer, {
    filename,
    contentType: file.mimetype || "application/octet-stream",
    knownLength: file.size,
  });
  formData.append("customPath", customPath);

  await axios.post(`${FILE_UPLOAD_BASE_URL}/upload-photo`, formData, {
    headers: formData.getHeaders(),
    maxBodyLength: Infinity,
  });

  const storedPath = `${customPath}${filename}`;
  const url = `${FILE_FETCH_BASE_URL}${encodeURIComponent(storedPath)}`;

  return { storedPath, url };
}

exports.getAllInsuranceRenewals = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const b = req.body || {};

    // ── Pagination ──────────────────────────────────────────
    const page = Math.max(1, parseInt(b.page) || 1);
    const pageSize = Math.max(1, parseInt(b.pageSize) || 10000);
    const offset = (page - 1) * pageSize;

    // ── Filters ─────────────────────────────────────────────
    const search = String(b.search || "").trim();
    const regNo = String(b.regNo || "")
      .trim()
      .toUpperCase();
    const insuType = String(b.insuType || "").trim();
    const paymentMode = String(b.paymentMode || "").trim();

    // ── Date helpers ─────────────────────────────────────────
    const pad2 = (x) => String(x).padStart(2, "0");

    const toDbDate = (v) => {
      if (!v) return null;
      const s = String(v).trim();

      let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;

      return null;
    };

    const toDisplay = (v) => {
      if (!v) return null;
      if (typeof v === "string") {
        const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return `${m[3]}/${m[2]}/${m[1]}`;
      }
      const d = v instanceof Date ? v : new Date(v);
      if (isNaN(d.getTime())) return null;
      return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
    };

    const policyFrom = toDbDate(b.policyFrom);
    const policyTo = toDbDate(b.policyTo);

    const createdFrom = toDbDate(b.createdFrom);
    const createdTo = toDbDate(b.createdTo);

    // ── WHERE ────────────────────────────────────────────────
    const whereParts = [];
    const replacements = {};

    // ✅ Only payment done rows
    whereParts.push(`P.PYMT_STATUS = 1`);

    if (search) {
      whereParts.push(`(
        IR.CUST_NAME       LIKE :search OR
        MST.VEHICAL_REG_NO LIKE :search OR
        IR.VEHICAL_REG_NO  LIKE :search OR
        IR.POLICY_NAME     LIKE :search OR
        CAST(IR.POLICY_NUMBER AS NVARCHAR) LIKE :search OR
        IR.MODEL_NAME      LIKE :search OR
        CAST(IR.CUST_MOB_NO AS NVARCHAR) LIKE :search
      )`);
      replacements.search = `%${search}%`;
    }
    console.log("whereParts", whereParts);

    if (regNo) {
      whereParts.push(`MST.VEHICAL_REG_NO = :regNo`);
      replacements.regNo = regNo;
    }

    if (insuType) {
      whereParts.push(`IR.INSU_TYPE = :insuType`);
      replacements.insuType = insuType;
    }

    if (paymentMode) {
      whereParts.push(`P.PYMT_MODE = :paymentMode`);
      replacements.paymentMode = paymentMode;
    }

    if (policyFrom && policyTo) {
      whereParts.push(`IR.POLICY_END_DATE BETWEEN :policyFrom AND :policyTo`);
      replacements.policyFrom = policyFrom;
      replacements.policyTo = policyTo;
    } else if (policyFrom) {
      whereParts.push(`IR.POLICY_END_DATE >= :policyFrom`);
      replacements.policyFrom = policyFrom;
    } else if (policyTo) {
      whereParts.push(`IR.POLICY_END_DATE <= :policyTo`);
      replacements.policyTo = policyTo;
    }

    if (createdFrom && createdTo) {
      whereParts.push(
        `CONVERT(date, P.CREATED_AT) BETWEEN :createdFrom AND :createdTo`,
      );
      replacements.createdFrom = createdFrom;
      replacements.createdTo = createdTo;
    } else if (createdFrom) {
      whereParts.push(`CONVERT(date, P.CREATED_AT) >= :createdFrom`);
      replacements.createdFrom = createdFrom;
    } else if (createdTo) {
      whereParts.push(`CONVERT(date, P.CREATED_AT) <= :createdTo`);
      replacements.createdTo = createdTo;
    }

    const whereSQL = whereParts.length
      ? `WHERE ${whereParts.join(" AND ")}`
      : "";

    // ── COUNT ────────────────────────────────────────────────
    const countSQL = `
      SELECT COUNT(*) AS total
      FROM (
        SELECT 
          P.UTD,
          ROW_NUMBER() OVER (
            PARTITION BY COALESCE(NULLIF(LTRIM(RTRIM(MST.VEHICAL_REG_NO)), ''), NULLIF(LTRIM(RTRIM(IR.VEHICAL_REG_NO)), ''), CAST(P.TRAN_ID AS VARCHAR))
            ORDER BY P.UTD DESC
          ) AS rn
        FROM dbo.INSU_RENEWAL_PYMT P
        LEFT JOIN dbo.INSU_RENEWAL_MST MST
          ON MST.UTD = P.TRAN_ID
        OUTER APPLY (
          SELECT TOP 1
            IR2.UTD,
            IR2.TRAN_ID,
            IR2.CUST_NAME,
            IR2.CUST_MOB_NO,
            IR2.VEHICAL_REG_NO,
            IR2.MODEL_NAME,
            IR2.POLICY_NAME,
            IR2.POLICY_NUMBER,
            IR2.POLICY_START_DATE,
            IR2.POLICY_END_DATE,
            IR2.INSU_TYPE,
            IR2.PREMIUM_AMOUNT
          FROM dbo.INSU_RENEWAL IR2
          WHERE IR2.TRAN_ID = MST.UTD
             OR (MST.VEHICAL_REG_NO IS NOT NULL AND REPLACE(UPPER(IR2.VEHICAL_REG_NO), ' ', '') = REPLACE(UPPER(MST.VEHICAL_REG_NO), ' ', ''))
             OR IR2.TRAN_ID = P.TRAN_ID
          ORDER BY IR2.UTD DESC
        ) IR
        ${whereSQL}
      ) Sub
      WHERE Sub.rn = 1
    `;

    console.log("sql", countSQL);

    const [countResult] = await sequelize.query(countSQL, {
      replacements,
      type: sequelize.QueryTypes.SELECT,
    });

    const totalRecords = parseInt(countResult?.total || 0);
    const totalPages = Math.ceil(totalRecords / pageSize) || 1;

    // ── DATA ─────────────────────────────────────────────────
    replacements.offset = offset;
    replacements.pageSize = pageSize;

    const dataSQL = `
      ;WITH FilteredPymt AS (
        SELECT
          P.UTD        AS PYMT_UTD,
          P.TRAN_ID    AS TRAN_ID,
          P.CREATED_AT AS PYMT_CREATED_AT,

          P.PYMT_MODE,
          COALESCE(PM_P.Misc_Name, PM_P.Misc_Abbr, P.PYMT_MODE, IR.PAYMENT_MODE) AS PAYMENT_MODE_NAME,
          P.PYMT_DATE,
          P.PYMT_AMOUNT,
          P.PYMT_CODE,
          COALESCE(P.PYMT_REMARK, P.REMARKS, IR.REMARKS) AS PYMT_REMARK,
          COALESCE(P.PYMT_REMARK, P.REMARKS, IR.REMARKS) AS REMARKS,
          P.PYMT_STATUS,

          P.ACNT_APPR_CODE,
          P.ACNT_APPR_REMARK,
          P.ACNT_APPR_STATUS,
          P.ACNT_APPR_DATE,

          P.UTR,
          P.BANK_NAME,
          P.CHEQUE_NO,
          P.DOC_PATH,

          P.PREMIUM_AMOUNT AS PYMT_PREMIUM_AMOUNT,

          MST.UTD            AS MST_UTD,
          MST.VEHICAL_REG_NO AS MST_REG_NO,

          IR.UTD            AS INSU_UTD,
          IR.CUST_NAME,
          IR.CUST_MOB_NO,
          IR.VEHICAL_REG_NO,
          IR.MODEL_NAME,
          IR.POLICY_NAME,
          IR.POLICY_NUMBER,
          IR.POLICY_START_DATE,
          IR.POLICY_END_DATE,
          IR.INSU_TYPE,
          IR.PREMIUM_AMOUNT AS IR_PREMIUM_AMOUNT,

          COALESCE(IR.PREMIUM_AMOUNT, P.PREMIUM_AMOUNT) AS PREMIUM_AMOUNT,

          -- Insurance Company Name: resolved from MISC_MST or fallback to POLICY_NAME
          COALESCE(
            NULLIF(NULLIF(IC_P.Misc_Name, ''), 'OTHER'),
            NULLIF(NULLIF(IC_IR.Misc_Name, ''), 'OTHER'),
            CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''), 'x')) = 1
                 THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                              WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                                AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(P.INSU_TYPE))
                              ORDER BY M.UTD), 'OTHER')
            END,
            CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(IR.INSU_TYPE)),''), 'x')) = 1
                 THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                              WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                                AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(IR.INSU_TYPE))
                              ORDER BY M.UTD), 'OTHER')
            END,
            CASE WHEN UPPER(LTRIM(RTRIM(P.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''),'x')) = 0
                 THEN LTRIM(RTRIM(P.INSU_TYPE))
            END,
            CASE WHEN UPPER(LTRIM(RTRIM(IR.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(IR.INSU_TYPE)),''),'x')) = 0
                 THEN LTRIM(RTRIM(IR.INSU_TYPE))
            END,
            NULLIF(IR.POLICY_NAME, ''),
            NULL
          ) AS INSU_COMPANY_NAME,

          COALESCE(
            NULLIF(NULLIF(IC_P.Misc_Name, ''), 'OTHER'),
            CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''), 'x')) = 1
                 THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                              WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                                AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(P.INSU_TYPE))
                              ORDER BY M.UTD), 'OTHER')
            END,
            CASE WHEN UPPER(LTRIM(RTRIM(P.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''),'x')) = 0
                 THEN LTRIM(RTRIM(P.INSU_TYPE))
            END,
            NULL
          ) AS PYMT_INSU_TYPE_NAME,

          DATEDIFF(day, CAST(GETDATE() AS date), IR.POLICY_END_DATE) AS DAYS_TO_EXPIRY,

          CASE
            WHEN IR.POLICY_END_DATE < CAST(GETDATE() AS date) THEN 'EXPIRED'
            WHEN IR.POLICY_END_DATE = CAST(GETDATE() AS date) THEN 'EXPIRING_TODAY'
            WHEN DATEDIFF(day, CAST(GETDATE() AS date), IR.POLICY_END_DATE) <= 30 THEN 'EXPIRING_SOON'
            ELSE 'ACTIVE'
          END AS EXPIRY_STATUS,

          ROW_NUMBER() OVER (
            PARTITION BY COALESCE(NULLIF(LTRIM(RTRIM(MST.VEHICAL_REG_NO)), ''), NULLIF(LTRIM(RTRIM(IR.VEHICAL_REG_NO)), ''), CAST(P.TRAN_ID AS VARCHAR))
            ORDER BY P.UTD DESC
          ) AS rn

        FROM dbo.INSU_RENEWAL_PYMT P
        LEFT JOIN dbo.INSU_RENEWAL_MST MST
          ON MST.UTD = P.TRAN_ID

        OUTER APPLY (
          SELECT TOP 1
            IR2.UTD,
            IR2.TRAN_ID,
            IR2.CUST_NAME,
            IR2.CUST_MOB_NO,
            IR2.VEHICAL_REG_NO,
            IR2.MODEL_NAME,
            IR2.POLICY_NAME,
            IR2.POLICY_NUMBER,
            IR2.POLICY_START_DATE,
            IR2.POLICY_END_DATE,
            IR2.INSU_TYPE,
            IR2.PREMIUM_AMOUNT,
            IR2.REMARKS,
            IR2.PAYMENT_MODE
          FROM dbo.INSU_RENEWAL IR2
          WHERE IR2.TRAN_ID = MST.UTD
             OR (MST.VEHICAL_REG_NO IS NOT NULL AND REPLACE(UPPER(IR2.VEHICAL_REG_NO), ' ', '') = REPLACE(UPPER(MST.VEHICAL_REG_NO), ' ', ''))
             OR IR2.TRAN_ID = P.TRAN_ID
          ORDER BY IR2.UTD DESC
        ) IR

        -- Insurance Company: match IR.INSU_TYPE by UTD, Misc_Code, Misc_Name, or Misc_Abbr
        LEFT JOIN dbo.MISC_MST IC_IR
          ON IC_IR.Misc_Type = 9
         AND ISNULL(IC_IR.Export_Type, 1) <> 33
         AND (
              CAST(IC_IR.UTD AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
           OR CAST(IC_IR.Misc_Code AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
           OR UPPER(LTRIM(RTRIM(IC_IR.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
           OR UPPER(LTRIM(RTRIM(IC_IR.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
         )

        -- Fallback: match P.INSU_TYPE
        LEFT JOIN dbo.MISC_MST IC_P
          ON IC_P.Misc_Type = 9
         AND ISNULL(IC_P.Export_Type, 1) <> 33
         AND (
              CAST(IC_P.UTD AS VARCHAR(50)) = CAST(P.INSU_TYPE AS VARCHAR(50))
           OR CAST(IC_P.Misc_Code AS VARCHAR(50)) = CAST(P.INSU_TYPE AS VARCHAR(50))
           OR UPPER(LTRIM(RTRIM(IC_P.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(P.INSU_TYPE AS nvarchar(200)))))
           OR UPPER(LTRIM(RTRIM(IC_P.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(P.INSU_TYPE AS nvarchar(200)))))
         )

        -- Payment Mode: match P.PYMT_MODE by UTD, Misc_Code, Misc_Name, or Misc_Abbr
        LEFT JOIN dbo.MISC_MST PM_P
          ON PM_P.Misc_Type = 18
         AND ISNULL(PM_P.Export_Type, 1) <> 33
         AND (
              CAST(PM_P.UTD AS VARCHAR(50)) = CAST(P.PYMT_MODE AS VARCHAR(50))
           OR CAST(PM_P.Misc_Code AS VARCHAR(50)) = CAST(P.PYMT_MODE AS VARCHAR(50))
           OR UPPER(LTRIM(RTRIM(PM_P.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(P.PYMT_MODE AS nvarchar(200)))))
           OR UPPER(LTRIM(RTRIM(PM_P.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(P.PYMT_MODE AS nvarchar(200)))))
         )

        -- Fallback Payment Mode: from IR.PAYMENT_MODE
        LEFT JOIN dbo.MISC_MST PM_IR
          ON PM_IR.Misc_Type = 18
         AND ISNULL(PM_IR.Export_Type, 1) <> 33
         AND (
              CAST(PM_IR.UTD AS VARCHAR(50)) = CAST(IR.PAYMENT_MODE AS VARCHAR(50))
           OR CAST(PM_IR.Misc_Code AS VARCHAR(50)) = CAST(IR.PAYMENT_MODE AS VARCHAR(50))
           OR UPPER(LTRIM(RTRIM(PM_IR.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(IR.PAYMENT_MODE AS nvarchar(200)))))
           OR UPPER(LTRIM(RTRIM(PM_IR.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(IR.PAYMENT_MODE AS nvarchar(200)))))
         )

        ${whereSQL}
      )
      SELECT *
      FROM FilteredPymt
      WHERE rn = 1
      ORDER BY PYMT_UTD DESC
      OFFSET :offset ROWS
      FETCH NEXT :pageSize ROWS ONLY
    `;

    const rows = await sequelize.query(dataSQL, {
      replacements,
      type: sequelize.QueryTypes.SELECT,
    });

    const data = rows.map((r) => ({
      UTD: r.PYMT_UTD,
      PYMT_UTD: r.PYMT_UTD,
      TRAN_ID: r.TRAN_ID,

      // Insurance Company Name (resolved from MISC_MST)
      INSU_COMPANY_NAME: r.INSU_COMPANY_NAME || null,

      CUST_NAME: r.CUST_NAME || null,
      CUST_MOB_NO: r.CUST_MOB_NO ? String(r.CUST_MOB_NO) : null,

      VEHICAL_REG_NO: r.MST_REG_NO || r.VEHICAL_REG_NO || null,
      MODEL_NAME: r.MODEL_NAME || null,

      POLICY_NAME: r.POLICY_NAME || null,
      POLICY_NUMBER: r.POLICY_NUMBER ? String(r.POLICY_NUMBER) : null,
      POLICY_START_DATE: toDisplay(r.POLICY_START_DATE),
      POLICY_END_DATE: toDisplay(r.POLICY_END_DATE),

      INSU_TYPE: r.PYMT_INSU_TYPE_NAME || r.INSU_COMPANY_NAME || null,
      PREMIUM_AMOUNT:
        r.PREMIUM_AMOUNT != null ? parseFloat(r.PREMIUM_AMOUNT) : null,

      PAYMENT_MODE: r.PAYMENT_MODE_NAME || null,
      PAYMENT_DATE: toDisplay(r.PYMT_DATE),
      PAYMENT_AMOUNT: r.PYMT_AMOUNT != null ? parseFloat(r.PYMT_AMOUNT) : null,

      UTR: r.UTR || null,
      CHEQUE_NO: r.CHEQUE_NO || null,
      BANK_NAME: r.BANK_NAME || null,
      DOC_PATH: r.DOC_PATH || null,
      CREATED_AT: toDisplay(r.PYMT_CREATED_AT),

      PYMT_CODE: r.PYMT_CODE || null,
      PYMT_REMARK: r.PYMT_REMARK || r.REMARKS || null,
      REMARKS: r.PYMT_REMARK || r.REMARKS || null,
      PYMT_STATUS: r.PYMT_STATUS != null ? Number(r.PYMT_STATUS) : null,
      ACNT_APPR_CODE: r.ACNT_APPR_CODE || null,
      ACNT_APPR_REMARK: r.ACNT_APPR_REMARK || null,
      ACNT_APPR_STATUS:
        r.ACNT_APPR_STATUS != null ? Number(r.ACNT_APPR_STATUS) : null,
      ACNT_APPR_DATE: toDisplay(r.ACNT_APPR_DATE),

      DAYS_TO_EXPIRY:
        r.DAYS_TO_EXPIRY != null ? parseInt(r.DAYS_TO_EXPIRY) : null,
      EXPIRY_STATUS: r.EXPIRY_STATUS || null,

      MST_UTD: r.MST_UTD || null,
      MST_REG_NO: r.MST_REG_NO || null,
    }));

    return res.json({
      success: true,
      Message: "Payments fetched (PYMT_STATUS=1).",
      pagination: {
        currentPage: page,
        pageSize,
        totalPages,
        totalRecords,
      },
      data,
    });
  } catch (e) {
    console.error("getAllInsuranceRenewals Error:", e);
    return res.status(500).json({ success: false, Message: e.message });
  } finally {
    await sequelize.close();
  }
};

exports.getInsuranceRenewalById = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const UTD = parseInt(req.body?.UTD ?? req.query?.UTD, 10);
    if (!Number.isFinite(UTD)) {
      return res
        .status(400)
        .json({ success: false, Message: "UTD required (integer)" });
    }

    // ── Date display helper ──────────────────────────────────
    const pad2 = (x) => String(x).padStart(2, "0");
    const toDisplay = (v) => {
      if (!v) return null;
      if (typeof v === "string") {
        const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return `${m[3]}/${m[2]}/${m[1]}`;
      }
      const d = v instanceof Date ? v : new Date(v);
      if (isNaN(d.getTime())) return null;
      return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
    };

    // ── Main detail query (BASE = INSU_RENEWAL_PYMT) ─────────
    const sql = `
      SELECT
        -- Payment identity
        P.UTD             AS PYMT_UTD,
        P.TRAN_ID         AS TRAN_ID,
        P.PYMT_STATUS     AS PYMT_STATUS,
        P.CREATED_AT      AS PYMT_CREATED_AT,

        -- Payment fields
        P.INSU_TYPE       AS PYMT_INSU_TYPE,
        P.PREMIUM_AMOUNT  AS PYMT_PREMIUM_AMOUNT,
        P.PYMT_MODE,
        COALESCE(PM_P.Misc_Name, PM_P.Misc_Abbr, P.PYMT_MODE, IR.PAYMENT_MODE) AS PAYMENT_MODE_NAME,
        P.PYMT_DATE,
        P.PYMT_AMOUNT,
        P.PYMT_CODE,
        COALESCE(P.PYMT_REMARK, P.REMARKS, IR.REMARKS) AS PYMT_REMARK,
        COALESCE(P.REMARKS, P.PYMT_REMARK, IR.REMARKS) AS REMARKS,
        P.UTR,
        P.CHEQUE_NO,
        P.BANK_NAME,
        P.DOC_PATH        AS PYMT_DOC_PATH,

        -- Accounts approval
        P.ACNT_APPR_CODE,
        P.ACNT_APPR_REMARK,
        P.ACNT_APPR_STATUS,
        P.ACNT_APPR_DATE,

        -- MST info
        MST.UTD            AS MST_UTD,
        MST.VEHICAL_REG_NO AS MST_REG_NO,
        MST.CREATED_AT     AS MST_CREATED_AT,
        MST.EXPORT_TYPE    AS MST_EXPORT_TYPE,

        -- Latest insurance detail via linking (MST -> INSU_RENEWAL)
        IR.UTD             AS INSU_UTD,
        IR.EXPORT_TYPE     AS EXPORT_TYPE,

        IR.CUST_NAME,
        IR.CUST_MOB_NO,

        IR.VEHICAL_REG_NO,
        IR.MODEL_NAME,

        IR.POLICY_NAME,
        IR.POLICY_NUMBER,
        IR.POLICY_START_DATE,
        IR.POLICY_END_DATE,

        IR.INSU_TYPE       AS IR_INSU_TYPE,
        IR.PREMIUM_AMOUNT  AS IR_PREMIUM_AMOUNT,

        -- Company Name
        COALESCE(
          NULLIF(NULLIF(IC_P.Misc_Name, ''), 'OTHER'),
          NULLIF(NULLIF(IC_IR.Misc_Name, ''), 'OTHER'),
          CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''), 'x')) = 1
               THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                            WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                              AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(P.INSU_TYPE))
                            ORDER BY M.UTD), 'OTHER')
          END,
          CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(IR.INSU_TYPE)),''), 'x')) = 1
               THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                            WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                              AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(IR.INSU_TYPE))
                            ORDER BY M.UTD), 'OTHER')
          END,
          CASE WHEN UPPER(LTRIM(RTRIM(P.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''),'x')) = 0
               THEN LTRIM(RTRIM(P.INSU_TYPE))
          END,
          CASE WHEN UPPER(LTRIM(RTRIM(IR.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(IR.INSU_TYPE)),''),'x')) = 0
               THEN LTRIM(RTRIM(IR.INSU_TYPE))
          END,
          NULLIF(IR.POLICY_NAME, ''),
          NULL
        ) AS INSU_COMPANY_NAME,
        COALESCE(
          NULLIF(NULLIF(IC_P.Misc_Name, ''), 'OTHER'),
          CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''), 'x')) = 1
               THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                            WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                              AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(P.INSU_TYPE))
                            ORDER BY M.UTD), 'OTHER')
          END,
          CASE WHEN UPPER(LTRIM(RTRIM(P.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''),'x')) = 0
               THEN LTRIM(RTRIM(P.INSU_TYPE))
          END,
          NULL
        ) AS PYMT_INSU_TYPE_NAME,

        -- Computed
        DATEDIFF(day, CAST(GETDATE() AS date), IR.POLICY_END_DATE) AS DAYS_TO_EXPIRY,
        CASE
          WHEN IR.POLICY_END_DATE < CAST(GETDATE() AS date) THEN 'EXPIRED'
          WHEN IR.POLICY_END_DATE = CAST(GETDATE() AS date) THEN 'EXPIRING_TODAY'
          WHEN DATEDIFF(day, CAST(GETDATE() AS date), IR.POLICY_END_DATE) <= 30 THEN 'EXPIRING_SOON'
          ELSE 'ACTIVE'
        END AS EXPIRY_STATUS

      FROM dbo.INSU_RENEWAL_PYMT P
      LEFT JOIN dbo.INSU_RENEWAL_MST MST
        ON MST.UTD = P.TRAN_ID
      OUTER APPLY (
        SELECT TOP 1
          IR2.UTD,
          IR2.TRAN_ID,
          IR2.EXPORT_TYPE,

          IR2.CUST_NAME,
          IR2.CUST_MOB_NO,

          IR2.VEHICAL_REG_NO,
          IR2.MODEL_NAME,

          IR2.POLICY_NAME,
          IR2.POLICY_NUMBER,
          IR2.POLICY_START_DATE,
          IR2.POLICY_END_DATE,

          IR2.INSU_TYPE,
          IR2.PREMIUM_AMOUNT,

          IR2.REMARKS,
          IR2.DOC_PATH,
          IR2.CREATED_AT,
          IR2.VALIDFROM,
          IR2.VALIDTO,
          IR2.PAYMENT_MODE
        FROM dbo.INSU_RENEWAL IR2
        WHERE IR2.TRAN_ID = MST.UTD
        ORDER BY IR2.UTD DESC
      ) IR

      LEFT JOIN dbo.MISC_MST IC_IR
        ON IC_IR.Misc_Type = 9
       AND ISNULL(IC_IR.Export_Type, 1) <> 33
       AND (
            CAST(IC_IR.UTD AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
         OR CAST(IC_IR.Misc_Code AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
         OR UPPER(LTRIM(RTRIM(IC_IR.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
         OR UPPER(LTRIM(RTRIM(IC_IR.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
       )

      LEFT JOIN dbo.MISC_MST IC_P
        ON IC_P.Misc_Type = 9
       AND ISNULL(IC_P.Export_Type, 1) <> 33
       AND (
            CAST(IC_P.UTD AS VARCHAR(50)) = CAST(P.INSU_TYPE AS VARCHAR(50))
         OR CAST(IC_P.Misc_Code AS VARCHAR(50)) = CAST(P.INSU_TYPE AS VARCHAR(50))
         OR UPPER(LTRIM(RTRIM(IC_P.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(P.INSU_TYPE AS nvarchar(200)))))
         OR UPPER(LTRIM(RTRIM(IC_P.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(P.INSU_TYPE AS nvarchar(200)))))
       )

      LEFT JOIN dbo.MISC_MST PM_P
        ON PM_P.Misc_Type = 18
       AND ISNULL(PM_P.Export_Type, 1) <> 33
       AND (
            CAST(PM_P.UTD AS VARCHAR(50)) = CAST(P.PYMT_MODE AS VARCHAR(50))
         OR CAST(PM_P.Misc_Code AS VARCHAR(50)) = CAST(P.PYMT_MODE AS VARCHAR(50))
         OR UPPER(LTRIM(RTRIM(PM_P.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(P.PYMT_MODE AS nvarchar(200)))))
         OR UPPER(LTRIM(RTRIM(PM_P.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(P.PYMT_MODE AS nvarchar(200)))))
       )

      LEFT JOIN dbo.MISC_MST PM_IR
        ON PM_IR.Misc_Type = 18
       AND ISNULL(PM_IR.Export_Type, 1) <> 33
       AND (
            CAST(PM_IR.UTD AS VARCHAR(50)) = CAST(IR.PAYMENT_MODE AS VARCHAR(50))
         OR CAST(PM_IR.Misc_Code AS VARCHAR(50)) = CAST(IR.PAYMENT_MODE AS VARCHAR(50))
         OR UPPER(LTRIM(RTRIM(PM_IR.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(IR.PAYMENT_MODE AS nvarchar(200)))))
         OR UPPER(LTRIM(RTRIM(PM_IR.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(IR.PAYMENT_MODE AS nvarchar(200)))))
       )
      WHERE P.UTD = :UTD
    `;

    const [row] = await sequelize.query(sql, {
      replacements: { UTD },
      type: Sequelize.QueryTypes.SELECT,
    });

    if (!row) {
      return res.status(404).json({
        success: false,
        Message: `No payment record found with UTD = ${UTD} in INSU_RENEWAL_PYMT`,
      });
    }

    // Vehicle reg no for history
    const regNoForHistory = row.MST_REG_NO || row.VEHICAL_REG_NO || null;

    // ── History (same vehicle ke previous insurance renewals) ─
    let history = [];
    if (regNoForHistory) {
      const histSQL = `
        SELECT
          IR.UTD,
          IR.POLICY_NAME,
          IR.POLICY_NUMBER,
          IR.POLICY_START_DATE,
          IR.POLICY_END_DATE,
          COALESCE(IC.Misc_Name, IC.Misc_Abbr, IR.INSU_TYPE, IR.POLICY_NAME) AS INSU_TYPE,
          IR.PREMIUM_AMOUNT,
          COALESCE(PM.Misc_Name, PM.Misc_Abbr, IR.PAYMENT_MODE) AS PAYMENT_MODE,
          IR.PAYMENT_AMOUNT,
          IR.PAYMENT_DATE,
          IR.DOC_PATH,
          IR.EXPORT_TYPE,
          IR.CREATED_AT
        FROM dbo.INSU_RENEWAL IR
        LEFT JOIN dbo.MISC_MST IC
          ON IC.Misc_Type = 9
         AND ISNULL(IC.Export_Type, 1) <> 33
         AND (
              CAST(IC.UTD AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
           OR CAST(IC.Misc_Code AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
           OR UPPER(LTRIM(RTRIM(IC.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
           OR UPPER(LTRIM(RTRIM(IC.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
         )
        LEFT JOIN dbo.MISC_MST PM
          ON PM.Misc_Type = 18
         AND ISNULL(PM.Export_Type, 1) <> 33
         AND (
              CAST(PM.UTD AS VARCHAR(50)) = CAST(IR.PAYMENT_MODE AS VARCHAR(50))
           OR CAST(PM.Misc_Code AS VARCHAR(50)) = CAST(IR.PAYMENT_MODE AS VARCHAR(50))
           OR UPPER(LTRIM(RTRIM(PM.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(IR.PAYMENT_MODE AS nvarchar(200)))))
           OR UPPER(LTRIM(RTRIM(PM.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(IR.PAYMENT_MODE AS nvarchar(200)))))
         )
        WHERE IR.VEHICAL_REG_NO = :regNo
        ORDER BY IR.UTD DESC
      `;

      const histRows = await sequelize.query(histSQL, {
        replacements: { regNo: regNoForHistory },
        type: Sequelize.QueryTypes.SELECT,
      });

      history = (histRows || []).map((h) => ({
        UTD: h.UTD,
        POLICY_NAME: h.POLICY_NAME || null,
        POLICY_NUMBER: h.POLICY_NUMBER ? String(h.POLICY_NUMBER) : null,
        POLICY_START_DATE: toDisplay(h.POLICY_START_DATE),
        POLICY_END_DATE: toDisplay(h.POLICY_END_DATE),
        INSU_TYPE: h.INSU_TYPE || null,
        PREMIUM_AMOUNT:
          h.PREMIUM_AMOUNT != null ? parseFloat(h.PREMIUM_AMOUNT) : null,
        PAYMENT_MODE: h.PAYMENT_MODE || null,
        PAYMENT_AMOUNT:
          h.PAYMENT_AMOUNT != null ? parseFloat(h.PAYMENT_AMOUNT) : null,
        PAYMENT_DATE: toDisplay(h.PAYMENT_DATE),
        DOC_PATH: h.DOC_PATH || null,
        EXPORT_TYPE: h.EXPORT_TYPE,
        CREATED_AT: toDisplay(h.CREATED_AT),
        IS_ACTIVE: h.EXPORT_TYPE === 1,
      }));
    }

    // ── Final response mapping ───────────────────────────────
    const data = {
      // ✅ Now main key is payment UTD
      UTD: row.PYMT_UTD,
      PYMT_UTD: row.PYMT_UTD,
      TRAN_ID: row.TRAN_ID || null,

      // Customer (linked from INSU_RENEWAL)
      CUST_NAME: row.CUST_NAME || null,
      CUST_MOB_NO: row.CUST_MOB_NO ? String(row.CUST_MOB_NO) : null,

      // Vehicle
      VEHICAL_REG_NO: row.MST_REG_NO || row.VEHICAL_REG_NO || null,
      MODEL_NAME: row.MODEL_NAME || null,

      // Policy
      POLICY_NAME: row.POLICY_NAME || null,
      POLICY_NUMBER: row.POLICY_NUMBER ? String(row.POLICY_NUMBER) : null,
      POLICY_START_DATE: toDisplay(row.POLICY_START_DATE),
      POLICY_END_DATE: toDisplay(row.POLICY_END_DATE),

      // Insurance Company
      INSU_COMPANY_NAME: row.INSU_COMPANY_NAME || row.IR_INSU_TYPE || row.PYMT_INSU_TYPE || null,
      INSU_TYPE: row.PYMT_INSU_TYPE_NAME || row.INSU_COMPANY_NAME || row.IR_INSU_TYPE || row.PYMT_INSU_TYPE || null,
      PREMIUM_AMOUNT:
        row.IR_PREMIUM_AMOUNT != null
          ? parseFloat(row.IR_PREMIUM_AMOUNT)
          : row.PYMT_PREMIUM_AMOUNT != null
            ? parseFloat(row.PYMT_PREMIUM_AMOUNT)
            : null,

      // Payment
      PAYMENT_MODE: row.PAYMENT_MODE_NAME || row.PYMT_MODE || null,
      PAYMENT_DATE: toDisplay(row.PYMT_DATE),
      PAYMENT_AMOUNT:
        row.PYMT_AMOUNT != null ? parseFloat(row.PYMT_AMOUNT) : null,
      UTR: row.UTR || null,
      CHEQUE_NO: row.CHEQUE_NO || null,
      BANK_NAME: row.BANK_NAME || null,

      // Payment meta
      PYMT_CODE: row.PYMT_CODE || null,
      PYMT_REMARK: row.PYMT_REMARK || row.REMARKS || null,
      REMARKS: row.PYMT_REMARK || row.REMARKS || null,
      PYMT_STATUS: row.PYMT_STATUS != null ? Number(row.PYMT_STATUS) : null,
      DOC_PATH: row.PYMT_DOC_PATH || null,
      CREATED_AT: toDisplay(row.PYMT_CREATED_AT),

      // Accounts approval
      ACNT_APPR_CODE: row.ACNT_APPR_CODE || null,
      ACNT_APPR_REMARK: row.ACNT_APPR_REMARK || null,
      ACNT_APPR_STATUS:
        row.ACNT_APPR_STATUS != null ? Number(row.ACNT_APPR_STATUS) : null,
      ACNT_APPR_DATE: toDisplay(row.ACNT_APPR_DATE),

      // Computed
      DAYS_TO_EXPIRY:
        row.DAYS_TO_EXPIRY != null ? parseInt(row.DAYS_TO_EXPIRY) : null,
      EXPIRY_STATUS: row.EXPIRY_STATUS || null,

      // MST
      master: {
        MST_UTD: row.MST_UTD || null,
        MST_REG_NO: row.MST_REG_NO || null,
        MST_CREATED_AT: toDisplay(row.MST_CREATED_AT),
        MST_EXPORT_TYPE: row.MST_EXPORT_TYPE || null,
      },

      history,
      historyCount: history.length,
    };

    return res.json({
      success: true,
      Message: "Record fetched successfully (by payment UTD)",
      data,
    });
  } catch (e) {
    console.error("getInsuranceRenewalById Error:", e);
    return res.status(500).json({ success: false, Message: e.message });
  } finally {
    await sequelize.close();
  }
};

exports.updateInsuranceRenewalApproval = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  const t = await sequelize.transaction();

  try {
    const InsuRenewalPymt = Insu_Renewal_Pymt(sequelize, DataTypes);
    const b = req.body || {};

    const UTD = parseInt(b.UTD ?? b.PYMT_UTD ?? req.query?.UTD, 10);
    if (!Number.isFinite(UTD)) {
      await t.rollback();
      return res.status(400).json({
        success: false,
        Message: "UTD required (payment UTD integer)",
      });
    }

    const ACTION = String(b.ACTION ?? b.status ?? "")
      .trim()
      .toUpperCase();
    if (!["APPROVE", "REJECT", "PENDING"].includes(ACTION)) {
      await t.rollback();
      return res.status(400).json({
        success: false,
        Message: "Invalid ACTION. Use APPROVE / REJECT / PENDING",
      });
    }

    const ACNT_APPR_REMARK = String(b.ACNT_APPR_REMARK ?? b.remark ?? "")
      .trim()
      .slice(0, 100);

    if (ACTION === "REJECT" && !ACNT_APPR_REMARK) {
      await t.rollback();
      return res.status(400).json({
        success: false,
        Message: "Reject ke liye ACNT_APPR_REMARK (remark) required hai",
      });
    }

    // ✅ resolve empcode (save-style + body support)
    const resolveEmpCode = async () => {
      // 1) direct empcode from req.user / headers / body
      let emp =
        req.user?.Emp_Code ??
        req.user?.EMPCODE ??
        req.user?.empcode ??
        req.headers?.empcode ??
        req.headers?.emp_code ??
        req.headers?.EMPCODE ??
        b.EmpCode ??
        b.empcode ??
        b.EMPCODE ??
        null;

      emp = emp != null ? String(emp).trim() : "";
      if (emp && emp !== "0") return emp;

      // 2) resolve using userCode/userName from req.user OR headers OR body
      const headerUserCode =
        req.headers?.usercode ??
        req.headers?.user_code ??
        req.headers?.userid ??
        req.headers?.user_id ??
        req.headers?.id ??
        null;

      const bodyUserCode =
        b.User_Code ?? b.usercode ?? b.user_code ?? b.UserCode ?? b.id ?? null;

      const userCode =
        req.user?.User_Code ??
        req.user?.UserCode ??
        headerUserCode ??
        bodyUserCode ??
        null;

      const headerUserName =
        req.headers?.name ??
        req.headers?.user_name ??
        req.headers?.username ??
        null;

      const bodyUserName =
        b.User_Name ??
        b.username ??
        b.user_name ??
        b.UserName ??
        b.name ??
        null;

      const userName =
        req.user?.User_Name ??
        req.user?.UserName ??
        headerUserName ??
        bodyUserName ??
        null;

      const userCodeNum =
        userCode != null &&
          String(userCode).trim() !== "" &&
          Number.isFinite(Number(userCode))
          ? Number(userCode)
          : null;

      const userNameStr =
        userName != null && String(userName).trim() !== ""
          ? String(userName).trim()
          : null;

      if (!userCodeNum && !userNameStr) return null;

      // ✅ Try 1: with Module_Code=10 (your system)
      let rows = await sequelize.query(
        `
        SELECT TOP 1 EMPCODE
        FROM dbo.user_tbl
        WHERE Export_Type = 1
          AND Module_Code = 10
          AND (
            (:userCode IS NOT NULL AND User_Code = :userCode)
            OR
            (:userName IS NOT NULL AND UPPER(LTRIM(RTRIM(User_Name))) = UPPER(LTRIM(RTRIM(:userName))))
          )
        ORDER BY User_Code DESC
        `,
        {
          type: Sequelize.QueryTypes.SELECT,
          replacements: { userCode: userCodeNum, userName: userNameStr },
          transaction: t,
        },
      );

      let dbEmp =
        rows?.[0]?.EMPCODE != null ? String(rows[0].EMPCODE).trim() : "";
      if (dbEmp && dbEmp !== "0") return dbEmp;

      // ✅ Try 2: without Module_Code filter (fallback)
      rows = await sequelize.query(
        `
        SELECT TOP 1 EMPCODE
        FROM dbo.user_tbl
        WHERE Export_Type = 1
          AND (
            (:userCode IS NOT NULL AND User_Code = :userCode)
            OR
            (:userName IS NOT NULL AND UPPER(LTRIM(RTRIM(User_Name))) = UPPER(LTRIM(RTRIM(:userName))))
          )
        ORDER BY User_Code DESC
        `,
        {
          type: Sequelize.QueryTypes.SELECT,
          replacements: { userCode: userCodeNum, userName: userNameStr },
          transaction: t,
        },
      );

      dbEmp = rows?.[0]?.EMPCODE != null ? String(rows[0].EMPCODE).trim() : "";
      if (dbEmp && dbEmp !== "0") return dbEmp;

      return null;
    };

    const empCode = await resolveEmpCode();

    if (!empCode && ACTION !== "PENDING") {
      await t.rollback();
      return res.status(401).json({
        success: false,
        Message:
          "EMPCODE not found for approver user. Send EmpCode/empcode in body OR send id/usercode/name headers/body.",
      });
    }

    // ensure row exists
    const exists = await InsuRenewalPymt.findOne({
      where: { UTD },
      raw: true,
      transaction: t,
    });

    if (!exists) {
      await t.rollback();
      return res.status(404).json({
        success: false,
        Message: `No payment record found with UTD=${UTD}`,
      });
    }

    const updateObj = {};

    if (ACTION === "APPROVE") {
      updateObj.ACNT_APPR_STATUS = 1;
      updateObj.ACNT_APPR_CODE = empCode;
      updateObj.ACNT_APPR_REMARK = ACNT_APPR_REMARK || null;
      updateObj.ACNT_APPR_DATE = Sequelize.literal("GETDATE()");
    } else if (ACTION === "REJECT") {
      updateObj.ACNT_APPR_STATUS = 0;
      updateObj.ACNT_APPR_CODE = empCode;
      updateObj.ACNT_APPR_REMARK = ACNT_APPR_REMARK || null;
      updateObj.ACNT_APPR_DATE = Sequelize.literal("GETDATE()");
    } else {
      updateObj.ACNT_APPR_STATUS = null;
      updateObj.ACNT_APPR_CODE = null;
      updateObj.ACNT_APPR_REMARK = null;
      updateObj.ACNT_APPR_DATE = null;
    }

    await InsuRenewalPymt.update(updateObj, {
      where: {
        [Op.or]: [
          { UTD },
          ...(exists.TRAN_ID ? [{ TRAN_ID: exists.TRAN_ID }] : []),
        ],
      },
      transaction: t,
    });

    await t.commit();

    return res.json({
      success: true,
      Message: `Updated approval status: ${ACTION}`,
      Result: {
        UTD,
        ACNT_APPR_STATUS: updateObj.ACNT_APPR_STATUS,
        ACNT_APPR_REMARK: updateObj.ACNT_APPR_REMARK,
        ACNT_APPR_CODE: updateObj.ACNT_APPR_CODE,
      },
    });
  } catch (e) {
    try {
      await t.rollback();
    } catch (_) { }
    return res.status(500).json({ success: false, Message: e.message });
  } finally {
    await sequelize.close();
  }
};

exports.getAllApprovedInsuranceRenewals = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const b = req.body || {};

    // ── Pagination ──────────────────────────────────────────
    const page = Math.max(1, parseInt(b.page) || 1);
    const pageSize = Math.max(1, parseInt(b.pageSize) || 10);
    const offset = (page - 1) * pageSize;

    // ── Filters ─────────────────────────────────────────────
    const search = String(b.search || "").trim();
    const regNo = String(b.regNo || "")
      .trim()
      .toUpperCase();
    const insuType = String(b.insuType || "").trim();
    const paymentMode = String(b.paymentMode || "").trim();

    // ✅ Employee codes
    const empcode = String(b.empcode || "").trim();
    const emp_dms_code = String(b.emp_dms_code || "").trim();

    // =========================================================
    // ✅ EMPLOYEE CODE REQUIRED
    // =========================================================
    if (!empcode) {
      return res.status(400).json({
        success: false,
        Message: "empcode required",
      });
    }

    // ✅ loc_code normalize
    const locRaw = b?.loc_code;
    let locCodes = [];

    if (Array.isArray(locRaw)) {
      locCodes = locRaw.map((x) => String(x).trim()).filter((s) => s !== "");
    } else if (
      locRaw !== undefined &&
      locRaw !== null &&
      String(locRaw).trim() !== ""
    ) {
      const s = String(locRaw).trim();
      if (s.includes(",")) {
        locCodes = s
          .split(",")
          .map((x) => String(x).trim())
          .filter((x) => x !== "");
      } else {
        locCodes = [s];
      }
    }

    console.log(locRaw);

    // =========================================================
    // ✅ ADMIN CHECK + EMPLOYEE FILTER
    // =========================================================
    let employeePlaceholders = "";
    let employeeReplacements = {};

    // अगर emp_dms_code है (Admin) तो सभी data
    const isAdmin =
      String(emp_dms_code || "").trim().toUpperCase() === "EDP";
    if (isAdmin) {
      console.log("✅ ADMIN MODE: Showing all approved payments");
      // Admin mode: कोई employee filter नहीं
      employeePlaceholders = "";
    } else {
      console.log("👤 EMPLOYEE MODE: Getting subordinates");
      // Normal employee: subordinates + खुद का data

      // STEP 1: Login employee के subordinates निकालो
      const reportingSql = `
        SELECT DISTINCT EMPCODE
        FROM EMPLOYEEMASTER
        WHERE (
          Reporting_1 = :empcode
          OR Reporting_2 = :empcode
          OR Reporting_3 = :empcode
        )
      `;

      const reportingEmps = await sequelize.query(reportingSql, {
        replacements: { empcode },
        type: sequelize.QueryTypes.SELECT,
      });

      // STEP 2: Employee codes को list में डालो
      let empCodeList = reportingEmps
        .map((e) => e.EMPCODE)
        .filter((code) => code !== null && code !== undefined)
        .map((code) => String(code).trim())
        .filter((code) => code !== "");

      // STEP 3: Login employee को खुद add करो
      if (!empCodeList.includes(String(empcode).trim())) {
        empCodeList.push(String(empcode).trim());
      }

      // STEP 4: Duplicates remove करो
      const uniqueEmpCodes = [...new Set(empCodeList)];

      console.log("Employee codes for filter:", uniqueEmpCodes);

      // STEP 5: Dynamic placeholders बनाओ
      employeePlaceholders = uniqueEmpCodes
        .map((_, index) => `:emp${index}`)
        .join(", ");

      uniqueEmpCodes.forEach((code, index) => {
        employeeReplacements[`emp${index}`] = code;
      });
    }

    // ✅ Branch filter -> MST UTD list
    let mstUtdList = null;
    if (locCodes.length) {
      const mstIds = await sequelize.query(
        `
          SELECT UTD
          FROM dbo.INSU_RENEWAL_MST
          WHERE LTRIM(RTRIM(CAST(loc_code AS VARCHAR(50)))) IN (:locCodes)
        `,
        {
          type: sequelize.QueryTypes.SELECT,
          replacements: { locCodes },
        }
      );

      console.log("getAllApprovedInsuranceRenewals: mstIds=", mstIds);

      mstUtdList = (mstIds || [])
        .map((r) => Number(r.UTD))
        .filter((n) => Number.isFinite(n));

      // no records for that branch => return empty
      if (!mstUtdList.length) {
        console.log(
          "getAllApprovedInsuranceRenewals: no MST UTD found for locCodes=",
          locCodes
        );

        return res.json({
          success: true,
          Message:
            "Approved payments fetched (PYMT_STATUS=1, ACNT_APPR_STATUS=1).",
          pagination: {
            currentPage: page,
            pageSize,
            totalPages: 1,
            totalRecords: 0,
          },
          data: [],
        });
      }
    }

    // ── Date helpers ─────────────────────────────────────────
    const pad2 = (x) => String(x).padStart(2, "0");

    const toDbDate = (v) => {
      if (!v) return null;
      const s = String(v).trim();

      let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;

      m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;

      return null;
    };

    const toDisplay = (v) => {
      if (!v) return null;
      if (typeof v === "string") {
        const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return `${m[3]}/${m[2]}/${m[1]}`;
      }
      const d = v instanceof Date ? v : new Date(v);
      if (isNaN(d.getTime())) return null;
      return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
    };

    const policyFrom = toDbDate(b.policyFrom);
    const policyTo = toDbDate(b.policyTo);

    // ✅ Custom date filter (based on ACNT_APPR_DATE)
    const apprFrom =
      toDbDate(b.createdFrom) ||
      toDbDate(b.createdAtFrom) ||
      toDbDate(b.createdDateFrom);

    const apprTo =
      toDbDate(b.createdTo) ||
      toDbDate(b.createdAtTo) ||
      toDbDate(b.createdDateTo);

    const apprOn =
      toDbDate(b.createdAt) || toDbDate(b.createdDate) || toDbDate(b.createdOn);

    // ── WHERE ────────────────────────────────────────────────
    const whereParts = [];
    const replacements = { ...employeeReplacements };

    // ✅ Only payment done rows
    whereParts.push(`P.PYMT_STATUS = 1`);

    // ✅ ONLY APPROVED rows
    whereParts.push(`ISNULL(P.ACNT_APPR_STATUS, 0) = 1`);

    // ✅ loc_code filter -> apply on P.TRAN_ID
    if (mstUtdList && mstUtdList.length) {
      whereParts.push(`CAST(P.TRAN_ID AS INT) IN (:mstUtdList)`);
      replacements.mstUtdList = mstUtdList;
    }

    // ✅ EMPLOYEE FILTER (अगर admin नहीं है)
    if (employeePlaceholders !== "") {
      whereParts.push(`P.PYMT_CODE IN (${employeePlaceholders})`);
    }

    if (search) {
      whereParts.push(`(
        IR.CUST_NAME       LIKE :search OR
        MST.VEHICAL_REG_NO LIKE :search OR
        IR.VEHICAL_REG_NO  LIKE :search OR
        IR.POLICY_NAME     LIKE :search OR
        CAST(IR.POLICY_NUMBER AS NVARCHAR) LIKE :search OR
        IR.MODEL_NAME      LIKE :search OR
        CAST(IR.CUST_MOB_NO AS NVARCHAR) LIKE :search
      )`);
      replacements.search = `%${search}%`;
    }

    if (regNo) {
      whereParts.push(`MST.VEHICAL_REG_NO = :regNo`);
      replacements.regNo = regNo;
    }

    if (insuType) {
      whereParts.push(`IR.INSU_TYPE = :insuType`);
      replacements.insuType = insuType;
    }

    if (paymentMode) {
      whereParts.push(`P.PYMT_MODE = :paymentMode`);
      replacements.paymentMode = paymentMode;
    }

    if (policyFrom && policyTo) {
      whereParts.push(`IR.POLICY_END_DATE BETWEEN :policyFrom AND :policyTo`);
      replacements.policyFrom = policyFrom;
      replacements.policyTo = policyTo;
    } else if (policyFrom) {
      whereParts.push(`IR.POLICY_END_DATE >= :policyFrom`);
      replacements.policyFrom = policyFrom;
    } else if (policyTo) {
      whereParts.push(`IR.POLICY_END_DATE <= :policyTo`);
      replacements.policyTo = policyTo;
    }

    // ✅ ACNT_APPR_DATE custom filter
    if (apprOn) {
      whereParts.push(`CONVERT(date, P.ACNT_APPR_DATE) = :apprOn`);
      replacements.apprOn = apprOn;
    } else if (apprFrom && apprTo) {
      whereParts.push(
        `CONVERT(date, P.ACNT_APPR_DATE) BETWEEN :apprFrom AND :apprTo`
      );
      replacements.apprFrom = apprFrom;
      replacements.apprTo = apprTo;
    } else if (apprFrom) {
      whereParts.push(`CONVERT(date, P.ACNT_APPR_DATE) >= :apprFrom`);
      replacements.apprFrom = apprFrom;
    } else if (apprTo) {
      whereParts.push(`CONVERT(date, P.ACNT_APPR_DATE) <= :apprTo`);
      replacements.apprTo = apprTo;
    }

    const whereSQL = whereParts.length
      ? `WHERE ${whereParts.join(" AND ")}`
      : "";

    // ── COUNT ────────────────────────────────────────────────
    const countSQL = `
      SELECT COUNT(*) AS total
      FROM dbo.INSU_RENEWAL_PYMT P
      LEFT JOIN dbo.INSU_RENEWAL_MST MST
        ON MST.UTD = P.TRAN_ID
      OUTER APPLY (
        SELECT TOP 1
          IR2.UTD,
          IR2.TRAN_ID,
          IR2.CUST_NAME,
          IR2.CUST_MOB_NO,
          IR2.VEHICAL_REG_NO,
          IR2.MODEL_NAME,
          IR2.POLICY_NAME,
          IR2.POLICY_NUMBER,
          IR2.POLICY_START_DATE,
          IR2.POLICY_END_DATE,
          IR2.INSU_TYPE,
          IR2.PREMIUM_AMOUNT
        FROM dbo.INSU_RENEWAL IR2
        WHERE IR2.TRAN_ID = MST.UTD
        ORDER BY IR2.UTD DESC
      ) IR
      ${whereSQL}
    `;

    const [countResult] = await sequelize.query(countSQL, {
      replacements,
      type: sequelize.QueryTypes.SELECT,
    });

    const totalRecords = parseInt(countResult?.total || 0, 10);
    const totalPages = Math.ceil(totalRecords / pageSize) || 1;

    // ── DATA ─────────────────────────────────────────────────
    replacements.offset = offset;
    replacements.pageSize = pageSize;

    const dataSQL = `
      SELECT
        P.UTD        AS PYMT_UTD,
        P.TRAN_ID    AS TRAN_ID,
        P.CREATED_AT AS PYMT_CREATED_AT,

        P.PYMT_MODE,
        P.PYMT_DATE,
        P.PYMT_AMOUNT,
        P.PYMT_CODE,
        P.PYMT_REMARK,
        P.PYMT_STATUS,

        P.ACNT_APPR_CODE,
        P.ACNT_APPR_REMARK,
        P.ACNT_APPR_STATUS,
        P.ACNT_APPR_DATE,

        P.UTR,
        P.BANK_NAME,
        P.REMARKS,
        P.DOC_PATH,

        P.PREMIUM_AMOUNT AS PYMT_PREMIUM_AMOUNT,

        MST.UTD            AS MST_UTD,
        MST.VEHICAL_REG_NO AS MST_REG_NO,
        MST.loc_code        AS MST_LOC_CODE,

        IR.UTD            AS INSU_UTD,
        IR.CUST_NAME,
        IR.CUST_MOB_NO,
        IR.VEHICAL_REG_NO,
        IR.MODEL_NAME,
        IR.POLICY_NAME,
        IR.POLICY_NUMBER,
        IR.POLICY_START_DATE,
        IR.POLICY_END_DATE,
        IR.INSU_TYPE,
        IR.PREMIUM_AMOUNT AS IR_PREMIUM_AMOUNT,

        COALESCE(IR.PREMIUM_AMOUNT, P.PREMIUM_AMOUNT) AS PREMIUM_AMOUNT,

        COALESCE(
          NULLIF(NULLIF(IC_P.Misc_Name, ''), 'OTHER'),
          NULLIF(NULLIF(IC_IR.Misc_Name, ''), 'OTHER'),
          CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''), 'x')) = 1
               THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                            WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                              AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(P.INSU_TYPE))
                            ORDER BY M.UTD), 'OTHER')
          END,
          CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(IR.INSU_TYPE)),''), 'x')) = 1
               THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                            WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                              AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(IR.INSU_TYPE))
                            ORDER BY M.UTD), 'OTHER')
          END,
          CASE WHEN UPPER(LTRIM(RTRIM(P.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''),'x')) = 0
               THEN LTRIM(RTRIM(P.INSU_TYPE))
          END,
          CASE WHEN UPPER(LTRIM(RTRIM(IR.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(IR.INSU_TYPE)),''),'x')) = 0
               THEN LTRIM(RTRIM(IR.INSU_TYPE))
          END,
          NULLIF(IR.POLICY_NAME, '')
        ) AS INSU_COMPANY_NAME,
        COALESCE(
          NULLIF(NULLIF(IC_P.Misc_Name, ''), 'OTHER'),
          CASE WHEN ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''), 'x')) = 1
               THEN NULLIF((SELECT TOP 1 M.Misc_Name FROM dbo.MISC_MST M
                            WHERE M.Misc_Type = 9 AND ISNULL(M.Export_Type,1)<>33
                              AND CAST(M.UTD AS NVARCHAR(20)) = LTRIM(RTRIM(P.INSU_TYPE))
                            ORDER BY M.UTD), 'OTHER')
          END,
          CASE WHEN UPPER(LTRIM(RTRIM(P.INSU_TYPE))) NOT IN ('OTHER','') AND ISNUMERIC(ISNULL(NULLIF(LTRIM(RTRIM(P.INSU_TYPE)),''),'x')) = 0
               THEN LTRIM(RTRIM(P.INSU_TYPE))
          END,
          NULL
        ) AS PYMT_INSU_TYPE_NAME,

        DATEDIFF(day, CAST(GETDATE() AS date), IR.POLICY_END_DATE) AS DAYS_TO_EXPIRY,

        CASE
          WHEN IR.POLICY_END_DATE < CAST(GETDATE() AS date) THEN 'EXPIRED'
          WHEN IR.POLICY_END_DATE = CAST(GETDATE() AS date) THEN 'EXPIRING_TODAY'
          WHEN DATEDIFF(day, CAST(GETDATE() AS date), IR.POLICY_END_DATE) <= 30 THEN 'EXPIRING_SOON'
          ELSE 'ACTIVE'
        END AS EXPIRY_STATUS

      FROM dbo.INSU_RENEWAL_PYMT P
      LEFT JOIN dbo.INSU_RENEWAL_MST MST
        ON MST.UTD = P.TRAN_ID

      OUTER APPLY (
        SELECT TOP 1
          IR2.UTD,
          IR2.TRAN_ID,
          IR2.CUST_NAME,
          IR2.CUST_MOB_NO,
          IR2.VEHICAL_REG_NO,
          IR2.MODEL_NAME,
          IR2.POLICY_NAME,
          IR2.POLICY_NUMBER,
          IR2.POLICY_START_DATE,
          IR2.POLICY_END_DATE,
          IR2.INSU_TYPE,
          IR2.PREMIUM_AMOUNT
        FROM dbo.INSU_RENEWAL IR2
        WHERE IR2.TRAN_ID = MST.UTD
        ORDER BY IR2.UTD DESC
      ) IR

      LEFT JOIN dbo.MISC_MST IC_IR
        ON IC_IR.Misc_Type = 9
       AND ISNULL(IC_IR.Export_Type, 1) <> 33
       AND (
            CAST(IC_IR.UTD AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
         OR CAST(IC_IR.Misc_Code AS VARCHAR(50)) = CAST(IR.INSU_TYPE AS VARCHAR(50))
         OR UPPER(LTRIM(RTRIM(IC_IR.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
         OR UPPER(LTRIM(RTRIM(IC_IR.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(IR.INSU_TYPE AS nvarchar(200)))))
       )

      LEFT JOIN dbo.MISC_MST IC_P
        ON IC_P.Misc_Type = 9
       AND ISNULL(IC_P.Export_Type, 1) <> 33
       AND (
            CAST(IC_P.UTD AS VARCHAR(50)) = CAST(P.INSU_TYPE AS VARCHAR(50))
         OR CAST(IC_P.Misc_Code AS VARCHAR(50)) = CAST(P.INSU_TYPE AS VARCHAR(50))
         OR UPPER(LTRIM(RTRIM(IC_P.Misc_Name))) = UPPER(LTRIM(RTRIM(CAST(P.INSU_TYPE AS nvarchar(200)))))
         OR UPPER(LTRIM(RTRIM(IC_P.Misc_Abbr))) = UPPER(LTRIM(RTRIM(CAST(P.INSU_TYPE AS nvarchar(200)))))
       )

      ${whereSQL}
      ORDER BY P.UTD DESC
      OFFSET :offset ROWS
      FETCH NEXT :pageSize ROWS ONLY
    `;

    const rows = await sequelize.query(dataSQL, {
      replacements,
      type: sequelize.QueryTypes.SELECT,
    });

    const data = (rows || []).map((r) => ({
      UTD: r.PYMT_UTD,
      PYMT_UTD: r.PYMT_UTD,
      TRAN_ID: r.TRAN_ID,

      INSU_COMPANY_NAME: r.INSU_COMPANY_NAME || null,

      CUST_NAME: r.CUST_NAME || null,
      CUST_MOB_NO: r.CUST_MOB_NO ? String(r.CUST_MOB_NO) : null,

      VEHICAL_REG_NO: r.MST_REG_NO || r.VEHICAL_REG_NO || null,
      MODEL_NAME: r.MODEL_NAME || null,

      POLICY_NAME: r.POLICY_NAME || null,
      POLICY_NUMBER: r.POLICY_NUMBER ? String(r.POLICY_NUMBER) : null,
      POLICY_START_DATE: toDisplay(r.POLICY_START_DATE),
      POLICY_END_DATE: toDisplay(r.POLICY_END_DATE),

      INSU_TYPE: r.PYMT_INSU_TYPE_NAME || r.INSU_COMPANY_NAME || null,
      PREMIUM_AMOUNT:
        r.PREMIUM_AMOUNT != null ? parseFloat(r.PREMIUM_AMOUNT) : null,

      PAYMENT_MODE: r.PYMT_MODE || null,
      PAYMENT_DATE: toDisplay(r.PYMT_DATE),
      PAYMENT_AMOUNT: r.PYMT_AMOUNT != null ? parseFloat(r.PYMT_AMOUNT) : null,

      UTR: r.UTR || null,
      BANK_NAME: r.BANK_NAME || null,
      DOC_PATH: r.DOC_PATH || null,
      CREATED_AT: toDisplay(r.PYMT_CREATED_AT),

      PYMT_CODE: r.PYMT_CODE || null,
      PYMT_REMARK: r.PYMT_REMARK || null,
      PYMT_STATUS: r.PYMT_STATUS != null ? Number(r.PYMT_STATUS) : null,

      ACNT_APPR_CODE: r.ACNT_APPR_CODE || null,
      ACNT_APPR_REMARK: r.ACNT_APPR_REMARK || null,
      ACNT_APPR_STATUS:
        r.ACNT_APPR_STATUS != null ? Number(r.ACNT_APPR_STATUS) : null,
      ACNT_APPR_DATE: toDisplay(r.ACNT_APPR_DATE),

      DAYS_TO_EXPIRY:
        r.DAYS_TO_EXPIRY != null ? parseInt(r.DAYS_TO_EXPIRY, 10) : null,
      EXPIRY_STATUS: r.EXPIRY_STATUS || null,

      MST_UTD: r.MST_UTD || null,
      MST_REG_NO: r.MST_REG_NO || null,
      LOC_CODE: r.MST_LOC_CODE || null,
    }));

    return res.json({
      success: true,
      Message: "Approved payments fetched (PYMT_STATUS=1, ACNT_APPR_STATUS=1).",
      pagination: {
        currentPage: page,
        pageSize,
        totalPages,
        totalRecords,
      },
      data,
      debug: {
        isAdmin: emp_dms_code && String(emp_dms_code).trim() !== "",
        employeeFiltered: employeePlaceholders !== "",
      },
    });
  } catch (e) {
    console.error("getAllApprovedInsuranceRenewals Error:", e);
    return res.status(500).json({ success: false, Message: e.message });
  } finally {
    await sequelize.close();
  }
};



exports.bonvoiceWebhook = async function bonvoiceWebhook(req, res) {
  let sequelize;

  const pick = (...vals) =>
    vals.find((v) => v !== undefined && v !== null && v !== "");

  try {
    // ✅ FIXED: Get compcode from query or default
    const compcode = "autovyn";
    req.headers.compcode = compcode;

    console.log("[BONVOICE] ====== WEBHOOK RECEIVED ======");
    console.log("[BONVOICE] Time:", new Date().toISOString());
    console.log("[BONVOICE] CompCode:", compcode);
    console.log("[BONVOICE] Headers:", {
      signature: req.headers["x-bonvoice-signature"] ? "present" : "missing",
      contentType: req.headers["content-type"],
    });

    console.log("[BONVOICE] ALL HEADERS:", req.headers);

    // ✅ STEP 2: Parse JSON body (raw Buffer supported)
    const rawBody = req.body;

    let body;
    if (Buffer.isBuffer(rawBody)) {
      try {
        body = JSON.parse(rawBody.toString("utf8"));
      } catch (e) {
        return res
          .status(400)
          .json({ Status: false, Message: "Invalid JSON body" });
      }
    } else {
      body = req.body || {};
    }

    // Normalize payload in case provider wraps with { data: {...} }
    const payload =
      body &&
        typeof body === "object" &&
        body.data &&
        typeof body.data === "object"
        ? body.data
        : body;

    // Debug logs (keep for now)
    const rawText = Buffer.isBuffer(req.body)
      ? req.body.toString("utf8")
      : JSON.stringify(req.body);
    console.log("[BONVOICE] RAW TEXT BODY:", rawText);
    console.log("[BONVOICE] PARSED BODY:", JSON.stringify(body, null, 2));
    console.log(
      "[BONVOICE] NORMALIZED PAYLOAD:",
      JSON.stringify(payload, null, 2)
    );

    // ✅ STEP 3: Webhook Verification (Optional)
    const enableVerification =
      String(process.env.BONVOICE_VERIFY || "").toLowerCase() === "true";
    let signatureVerified = false;

    if (enableVerification) {
      console.log("[BONVOICE] Verification ENABLED");

      const signature =
        req.headers["x-bonvoice-signature"] ||
        req.headers["x-webhook-signature"] ||
        req.headers["x-signature"] ||
        null;

      const secret =
        process.env.BONVOICE_WEBHOOK_SECRET ||
        "663e4efc73b2fe6d11e8ed0df6e5a5952dae40a6e04393e9";

      if (!secret) {
        return res.status(500).json({
          Status: false,
          Message: "BONVOICE_WEBHOOK_SECRET missing in env",
        });
      }

      if (!signature) {
        return res.status(401).json({
          Status: false,
          Message: "Unauthorized - No signature provided",
        });
      }

      if (!Buffer.isBuffer(rawBody)) {
        return res.status(500).json({
          Status: false,
          Message:
            "Webhook misconfigured. Use express.raw({type:'application/json'}) on this route.",
        });
      }

      signatureVerified = verifyWebhook(rawBody, signature, secret);
      if (!signatureVerified) {
        return res.status(401).json({
          Status: false,
          Message: "Unauthorized - Invalid signature",
        });
      }

      console.log("[BONVOICE] Signature verified successfully");
    } else {
      console.log("[BONVOICE] Verification DISABLED");
    }

    // ✅ STEP 4: Event filter (only call.completed)
    const eventName = pick(payload?.event, body?.event);

    if (eventName && eventName !== "call.completed") {
      console.log("[BONVOICE] Event ignored:", eventName);
      return res
        .status(200)
        .json({ Status: true, Message: "Event ignored", event: eventName });
    }

    // ✅ STEP 5: DB connect
    sequelize = await dbname(req, compcode);
    console.log("[BONVOICE] Database connected compcode:", compcode);

    // ✅ STEP 6A: Find callId (works for both full payload and {callId})
    const callId = pick(
      payload?.call?.id,
      payload?.call?.callId,
      payload?.callId,
      payload?.call_id,
      body?.callId,
      body?.call_id
    );

    if (!callId) {
      return res
        .status(400)
        .json({ Status: false, Message: "call_id missing" });
    }

    // ✅ STEP 6B: If payload is ONLY {callId}, enrich using mapping + GET /leads/{leadId}
    let effectivePayload = payload;

    const isCallIdOnly =
      effectivePayload &&
      typeof effectivePayload === "object" &&
      effectivePayload.callId &&
      !effectivePayload.event &&
      !effectivePayload.lead &&
      !effectivePayload.call;

    if (isCallIdOnly) {
      console.log(
        "[BONVOICE] callId-only webhook detected. Enrichment will run.",
        {
          callId: String(callId),
        }
      );

      // 1) Get lead_id from your DB mapping table (call_webhook_dtl)
      const { QueryTypes } = require("sequelize");
      const mapRows = await sequelize.query(
      `SELECT TOP 1
    lead_id,
    phone_number,
    callee_name,
    campaign_id
FROM dbo.call_webhook_dtl
WHERE call_id = :callId
ORDER BY created_at DESC`,
        {
          replacements: { callId: String(callId).trim() },
          type: QueryTypes.SELECT,
        }
      );

      const mappedLeadId = mapRows?.[0]?.lead_id
        ? String(mapRows[0].lead_id).trim()
        : null;

      if (!mappedLeadId) {
        console.warn(
          "[BONVOICE] No lead_id mapping found for callId. Saving minimal data only.",
          {
            callId: String(callId),
          }
        );

        // Keep effectivePayload as-is
      } else {
        // 2) Fetch lead + latestCall from Bonvoice (if client is configured)
        try {
          // Note: Requires axios or similar HTTP client configured as 'client'
          const leadRes = await client.get(`/leads/${mappedLeadId}`);
          const leadFull = leadRes.data || {};

          // 3) Convert into webhook-like structure
          effectivePayload = {
            event: "call.completed",
            timestamp: new Date().toISOString(),
            lead: {
              id: leadFull?.id,
              name: leadFull?.name,
              phone: leadFull?.phone,
              email: leadFull?.email,
              status: leadFull?.leadStatus,
              column1: leadFull?.column1,
              column2: leadFull?.column2,
              column3: leadFull?.column3,
              classification: leadFull?.classification,
            },
            call: {
              id: leadFull?.latestCall?.id,
              status: leadFull?.latestCall?.status,
              duration_seconds: leadFull?.latestCall?.duration_seconds,
              summary: leadFull?.latestCall?.summary,
              transcript: leadFull?.latestCall?.transcript,
              recording_url: leadFull?.latestCall?.recording_url,
            },
          };

          // Safety check
          if (
            String(effectivePayload?.call?.id || "") !== String(callId)
          ) {
            console.warn(
              "[BONVOICE] latestCall.id does not match webhook callId. Will not overwrite call fields.",
              {
                webhookCallId: String(callId),
                latestCallId: String(effectivePayload?.call?.id || ""),
              }
            );

            effectivePayload.call = { id: callId };
          }

          console.log("[BONVOICE] Enrichment done using /leads/{leadId}", {
            mappedLeadId,
            callId: String(callId),
          });
        } catch (enrichErr) {
          console.warn(
            "[BONVOICE] Enrichment API failed:",
            enrichErr.message
          );
          // Continue with minimal data
        }
      }
    }

    // ✅ STEP 6C: Extract from effectivePayload
    const lead = effectivePayload?.lead || {};
    const call = effectivePayload?.call || {};

    const leadId = pick(
      lead?.id,
      lead?.leadId,
      lead?.lead_id,
      effectivePayload?.leadId,
      effectivePayload?.lead_id
    );

    const campaignId =
      pick(call?.prompt_name, call?.promptName, effectivePayload?.promptName) ||
      null;

    const status =
      pick(call?.status, effectivePayload?.status, body?.status) || "completed";

    const phoneNumber =
      pick(lead?.phone, lead?.phoneNumber, lead?.mobile) || null;

    const durationSeconds =
      pick(call?.duration_seconds, call?.durationSeconds, call?.duration) ||
      null;

    const recordingUrl = pick(call?.recording_url, call?.recordingUrl) || null;

    const triggeredAt =
      pick(effectivePayload?.timestamp, body?.timestamp, call?.timestamp) ||
      null;

    const data = {
      call_id: String(callId).trim(),
      lead_id: leadId ? String(leadId).trim() : null,
      campaign_id: campaignId ? String(campaignId).substring(0, 100) : null,
      direction: "outbound",
      phone_number: phoneNumber ? String(phoneNumber).substring(0, 15) : null,
      callee_name: lead?.name ? String(lead.name).substring(0, 100) : null,
      status: String(status).substring(0, 30),
      duration:
        durationSeconds !== undefined &&
          durationSeconds !== null &&
          durationSeconds !== ""
          ? Number(durationSeconds)
          : null,
      summary: pick(call?.summary, effectivePayload?.summary) || null,
      transcript: pick(call?.transcript, effectivePayload?.transcript) || null,
      category:
        pick(call?.classification, lead?.classification, call?.category) ||
        null,
      recording_url: recordingUrl
        ? String(recordingUrl).substring(0, 500)
        : null,
      triggered_at: triggeredAt,
      start_time: pick(call?.started_at, call?.start_time) || null,
      end_time: pick(call?.ended_at, call?.end_time, triggeredAt) || null,
      column1: pick(lead?.column1) || null,
      column2: pick(lead?.column2) || null,
      column3: pick(lead?.column3) || null,
    };

    console.log("[BONVOICE] Extracted(normalized):", {
      event: pick(effectivePayload?.event, eventName),
      call_id: data.call_id,
      lead_id: data.lead_id,
      phone: data.phone_number,
      name: data.callee_name,
      duration: data.duration,
      recording_url: data.recording_url,
      status: data.status,
      category: data.category,
    });

    // ✅ STEP 7: MERGE into call_webhook_dtl (WITHOUT lead_id update)
    const mergeSql = `
      MERGE dbo.call_webhook_dtl AS T
      USING (SELECT :call_id AS call_id) AS S
      ON T.call_id = S.call_id
      WHEN MATCHED THEN
        UPDATE SET
          campaign_id   = ISNULL(:campaign_id, T.campaign_id),
          direction     = ISNULL(:direction, T.direction),
          phone_number  = ISNULL(:phone_number, T.phone_number),
          status        = ISNULL(:status, T.status),
          duration      = ISNULL(:duration, T.duration),
          summary       = ISNULL(:summary, T.summary),
          transcript    = ISNULL(:transcript, T.transcript),
          callee_name   = ISNULL(:callee_name, T.callee_name),
          category      = ISNULL(:category, T.category),
          recording_url = ISNULL(:recording_url, T.recording_url),
          triggered_at  = ISNULL(:triggered_at, T.triggered_at),
          start_time    = ISNULL(:start_time, T.start_time),
          end_time      = ISNULL(:end_time, T.end_time)
      WHEN NOT MATCHED THEN
        INSERT (call_id, campaign_id, direction, phone_number, status, duration,
                summary, transcript, callee_name, category, recording_url,
                triggered_at, start_time, end_time, created_at)
        VALUES (:call_id, :campaign_id, :direction, :phone_number, :status, :duration,
                :summary, :transcript, :callee_name, :category, :recording_url,
                :triggered_at, :start_time, :end_time, GETDATE());
    `;

    await sequelize.query(mergeSql, {
      replacements: {
        call_id: data.call_id,
        campaign_id: data.campaign_id || null,
        direction: data.direction || null,
        phone_number: data.phone_number || null,
        status: data.status || null,
        duration: data.duration,
        summary: data.summary ? String(data.summary).substring(0, 8000) : null,
        transcript: data.transcript
          ? String(data.transcript).substring(0, 8000)
          : null,
        callee_name: data.callee_name
          ? String(data.callee_name).substring(0, 100)
          : null,
        category: data.category
          ? String(data.category).substring(0, 100)
          : null,
        recording_url: data.recording_url
          ? String(data.recording_url).substring(0, 500)
          : null,
        triggered_at: data.triggered_at || null,
        start_time: data.start_time || null,
        end_time: data.end_time || null,
      },
    });

    console.log("[BONVOICE] call_webhook_dtl MERGE completed");

    let tranId = null;

    if (data.lead_id) {
      const { QueryTypes } = require("sequelize");
      const followupRows = await sequelize.query(
        `SELECT TOP 1 lead_id, TRAN_ID
         FROM dbo.FOLLOWUP_DETAILS
         WHERE bonvoice_lead_id = :leadId
         ORDER BY CREATED_AT DESC`,
        {
          replacements: { leadId: data.lead_id },
          type: QueryTypes.SELECT,
        }
      );

      tranId = followupRows?.[0]?.TRAN_ID
        ? String(followupRows[0].TRAN_ID).trim()
        : null;
    }

    if (tranId) {
      const followupStatus =
        data.category && String(data.category).toLowerCase() === "hot"
          ? "INTERESTED"
          : "CALL_COMPLETED";

      try {
        const { QueryTypes } = require("sequelize");
        await sequelize.query(
          `UPDATE dbo.FOLLOWUP_DETAILS
           SET
             FOLLOWUP_STATUS = :status,
             FOLLOWUP_DATE = GETDATE(),
             BONVOICE_LEAD_ID = :leadId
           WHERE TRAN_ID = :tranId`,
          {
            replacements: {
              status: followupStatus,
              tranId,
              leadId: data.lead_id || null,
            },
            type: QueryTypes.UPDATE,
          }
        );

        console.log("[BONVOICE] FOLLOWUP_DETAILS updated:", {
          TRAN_ID: tranId,
          bonvoice_lead_id: data.lead_id,
          status: followupStatus,
        });
      } catch (fuErr) {
        console.warn(
          "[BONVOICE] FOLLOWUP_DETAILS update failed:",
          fuErr.message
        );
      }
    } else {
      console.log("[BONVOICE] No TRAN_ID found for lead_id:", data.lead_id);
    }

    console.log("[BONVOICE] ====== WEBHOOK PROCESSED SUCCESSFULLY ======");

    return res.status(200).json({
      Status: true,
      Message: "Webhook processed - data saved",
      verified: enableVerification ? signatureVerified : false,
      callId: data.call_id,
      leadId: data.lead_id,
      tranId: tranId || null,
      saved: true,
      data_saved: {
        call_id: data.call_id,
        lead_id: data.lead_id,
        phone_number: data.phone_number,
        callee_name: data.callee_name,
        status: data.status,
        duration: data.duration,
        category: data.category,
        recording_url: data.recording_url,
        followup_tran_id: tranId || null,
      },
    });
  } catch (err) {
    console.error("[BONVOICE] ERROR:", err.message);
    console.error("[BONVOICE] STACK:", err.stack);
    return res.status(500).json({
      Status: false,
      Message: err.message,
      error: err?.original?.message,
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};

// function extractErrorDetails(err) {
//   const details = {
//     message: err?.message || "Unknown error",
//     original: err?.original?.message || null,
//     number: err?.original?.number || null,
//     lineNumber: err?.original?.lineNumber || null,
//     state: err?.original?.state || null,
//     class: err?.original?.class || null,
//   };

//   // Extract from Tedious errors
//   if (err?.parent?.errors && Array.isArray(err.parent.errors)) {
//     details.tedious_errors = err.parent.errors.map((e) => ({
//       message: e?.message || String(e),
//       number: e?.number,
//       state: e?.state,
//       lineNumber: e?.lineNumber,
//     }));
//   }

//   return details;
// }

function normalizePhone(p) {
  if (!p) return "";
  p = String(p).trim().replace(/[\s-]/g, "");
  if (/^\d{10}$/.test(p)) return `+91${p}`;
  if (/^91\d{10}$/.test(p)) return `+${p}`;
  return p;
}

function cleanVehicleNo(v) {
  return String(v || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  const parts = String(dateStr).split("-");
  if (parts.length !== 3) return String(dateStr);
  const [y, m, d] = parts;
  return `${d.padStart(2, "0")}/${m.padStart(2, "0")}/${y}`;
}

const pickVal = (...vals) =>
  vals.find((v) => v !== undefined && v !== null && v !== "");

const toJsonStringHelper = (v) => {
  if (v === undefined || v === null) return null;
  if (typeof v === "string") return v;
  try {
    return JSON.stringify(v);
  } catch {
    return String(v);
  }
};

const toSqlDateTimeOrNullHelper = (v) => {
  if (!v) return null;
  const d = typeof v === "number" ? (v > 1e11 ? new Date(v) : new Date(v * 1000)) : new Date(v);
  if (isNaN(d.getTime())) return null;

  // Convert to IST (UTC + 5:30)
  const istOffset = 5.5 * 60 * 60 * 1000;
  const istDate = new Date(d.getTime() + (d.getTimezoneOffset() * 60 * 1000) + istOffset);

  const pad2 = (n) => String(n).padStart(2, "0");
  const pad3 = (n) => String(n).padStart(3, "0");

  return (
    `${istDate.getFullYear()}-${pad2(istDate.getMonth() + 1)}-${pad2(istDate.getDate())} ` +
    `${pad2(istDate.getHours())}:${pad2(istDate.getMinutes())}:${pad2(istDate.getSeconds())}.` +
    `${pad3(istDate.getMilliseconds())}`
  );
};

const webhookSaveLocks = new Map();

const extractCleanSummary = (val) => {
  if (!val) return null;

  // If val is an array (Callmatic insights array)
  if (Array.isArray(val)) {
    // 1. Look for actionType: SUMMARIZE
    const sumItem = val.find(
      (x) =>
        x &&
        (String(x.actionType || "").toUpperCase() === "SUMMARIZE" ||
          String(x.type || "").toUpperCase() === "SUMMARIZE" ||
          x.output?.summary ||
          x.summary)
    );
    if (sumItem) {
      const res =
        sumItem.output?.summary ||
        sumItem.summary ||
        sumItem.text ||
        sumItem.output?.text ||
        sumItem.output?.overview ||
        sumItem.overview;
      if (res && typeof res === "string" && res.trim().length > 0)
        return res.trim();
    }
    // 2. Any item that has output.summary or summary
    for (const item of val) {
      if (item && typeof item === "object") {
        const s =
          item.output?.summary ||
          item.summary ||
          item.overview ||
          item.call_summary ||
          item.output?.text;
        if (s && typeof s === "string" && s.trim().length > 0) return s.trim();
      }
    }
    return null;
  }

  // If val is an object
  if (typeof val === "object") {
    if (String(val.actionType || "").toUpperCase() === "SUMMARIZE") {
      const s =
        val.output?.summary ||
        val.summary ||
        val.output?.text ||
        val.text ||
        val.overview;
      if (s && typeof s === "string" && s.trim().length > 0) return s.trim();
    }
    const inner =
      val.summary ||
      val.output?.summary ||
      val.call_summary ||
      val.overview ||
      val.text ||
      val.description;
    if (typeof inner === "string" && inner.trim().length > 0)
      return inner.trim();

    if (val.insights) {
      const fromIns = extractCleanSummary(val.insights);
      if (fromIns) return fromIns;
    }
    return null;
  }

  // If val is a string
  if (typeof val === "string") {
    const s = val.trim();
    if (!s) return null;

    // Check if it's JSON array or object
    if (s.startsWith("[") || s.startsWith("{")) {
      try {
        const parsed = JSON.parse(s);
        const res = extractCleanSummary(parsed);
        if (res) return res;
      } catch (_) { }
    }

    // Try regex search for SUMMARIZE output summary
    const sumMatch =
      s.match(
        /"actionType"\s*:\s*"SUMMARIZE"[^}]*"output"\s*:\s*\{[^}]*"summary"\s*:\s*"([^"\\]*(?:\\.[^"\\]*)*)"/i
      ) ||
      s.match(
        /"actionType"\s*:\s*"SUMMARIZE"[^}]*"summary"\s*:\s*"([^"\\]*(?:\\.[^"\\]*)*)"/i
      ) ||
      s.match(
        /"summary"\s*:\s*"([^"\\]*(?:\\.[^"\\]*)*)"/i
      );

    if (sumMatch && sumMatch[1]) {
      try {
        return JSON.parse(`"${sumMatch[1]}"`);
      } catch {
        return sumMatch[1];
      }
    }

    // If it is a raw JSON string like {"actionType"..., don't return raw JSON string
    if (s.startsWith('{"actionType"') || s.startsWith('[{"actionType"')) {
      return null;
    }

    return s;
  }

  return String(val);
};

const extractSummaryHelper = (data = {}) => {
  if (!data) return null;

  // 1. Direct summary fields
  const direct =
    data.summary ||
    data.SUMMARY ||
    data.call_summary ||
    data.transcript_summary ||
    data.ai_summary ||
    data.call?.summary ||
    data.call?.call_summary ||
    data.call?.ai_summary ||
    data.call?.transcript_summary ||
    data.analysis?.summary ||
    data.analysis?.call_summary ||
    data.analysis?.overview ||
    data.analysis?.structured_data?.summary ||
    data.analysis?.transcript_summary ||
    data.call_analysis?.summary ||
    data.call_analysis?.call_summary ||
    data.call_analysis?.overview ||
    data.overview ||
    data.notes ||
    data.data?.summary ||
    data.data?.call_summary ||
    data.data?.analysis?.summary ||
    data.data?.analysis?.call_summary ||
    data.data?.analysis?.overview ||
    data.data?.call_analysis?.summary ||
    data.data?.call_analysis?.call_summary ||
    data.data?.ai_summary;

  const fromDirect = extractCleanSummary(direct);
  if (fromDirect) return fromDirect;

  // 2. Check insights field (Callmatic specific)
  if (data.insights !== undefined && data.insights !== null) {
    const fromInsights = extractCleanSummary(data.insights);
    if (fromInsights) return fromInsights;
  }

  // 3. Check nested insights
  if (data.data?.insights || data.call?.insights) {
    const fromNested = extractCleanSummary(data.data?.insights || data.call?.insights);
    if (fromNested) return fromNested;
  }

  return null;
};

const saveCallWebhookDetails = async (sequelize, callId, data = {}) => {
  if (!sequelize || !callId) return null;
  const cid = String(callId).trim();
  if (!cid) return null;

  // In-memory Mutex queue per call_id so parallel writes for same call wait in sequence
  let previousLock = webhookSaveLocks.get(cid) || Promise.resolve();
  let currentLockResolver = () => { };
  const currentLock = new Promise((resolve) => {
    currentLockResolver = resolve;
  });
  webhookSaveLocks.set(cid, currentLock);

  try {
    await previousLock;

    const campaign_id = pickVal(data.campaignId, data.campaign_id, data.campaign) || null;
    const direction = pickVal(data.direction) || "outbound";
    const phone_number = pickVal(
      data.phoneNumber,
      data.phone_number,
      data.phone,
      data.lead?.phone,
      data.lead?.phone_number,
      data.to
    ) || null;

    let statusRaw = pickVal(
      data.status,
      data.callStatus,
      data.call_status,
      data.call?.status,
      data.call?.callStatus,
      data.data?.status,         // 👈 Callmatic nested status
      data.data?.callStatus,
      data.state
    );
    let status = statusRaw ? String(statusRaw).toLowerCase().trim() : null;
    // 2. Agar status nahi mila lekin event 'call.completed' ya 'call.ended' aaya hai
    if (!status && (data.event || data.data?.event)) {
      const ev = String(data.event || data.data?.event).toLowerCase();
      if (ev.includes("complete") || ev.includes("end") || ev.includes("finish")) {
        status = "completed";
      } else if (ev.includes("fail") || ev.includes("error")) {
        status = "failed";
      }
    }
    // 3. 'ended' ya 'finished' ko standard 'completed' me convert karo
    if (status && ["ended", "finished", "success", "call_ended", "call.completed", "call.ended"].includes(status)) {
      status = "completed";
    }
    // 4. 🔥 FAILSAFE: Agar transcript ya recording_url aa gayi hai, to call 100% complete ho chuki hai
    if (
      (!status || status === "initiated" || status === "ringing" || status === "pending") &&
      (data.transcript || data.data?.transcript || data.recording_url || data.data?.recording_url || data.recording)
    ) {
      status = "completed";
    }

    const triggered_at = toSqlDateTimeOrNullHelper(
      pickVal(data.triggeredAt, data.triggered_at, data.created_at, data.createdAt)
    );
    const start_time = toSqlDateTimeOrNullHelper(
      pickVal(data.startTime, data.start_time, data.startedAt, data.started_at)
    );
    const end_time = toSqlDateTimeOrNullHelper(
      pickVal(data.endTime, data.end_time, data.endedAt, data.ended_at)
    );

    const duration = pickVal(data.duration, data.duration_seconds, data.call?.duration_seconds, data.call?.duration) ?? null;
    const ring_time = pickVal(data.ringTime, data.ring_time) ?? null;

    const summary = extractSummaryHelper(data);
    console.log(`[SAVE CALL WEBHOOK DETAILS] CallID: ${cid} | Extracted Summary:`, summary);

    const transcriptRaw = pickVal(data.transcript, data.call?.transcript, data.data?.transcript) ?? null;
    const transcript = toJsonStringHelper(transcriptRaw);

    const callee_name = pickVal(
      data.callee_name,
      data.calleeName,
      data.lead?.name,
      data.variables?.callee_name
    ) ?? null;
    const category = pickVal(data.category, data.classification, data.call?.classification, data.lead?.classification) ?? null;
    const recording_url = pickVal(
      data.recording_url,
      data.recordingUrl,
      data.recording,
      data.call?.recording_url
    ) ?? null;

    const upsertSql = `
      IF EXISTS (SELECT 1 FROM dbo.call_webhook_dtl WITH (UPDLOCK, HOLDLOCK) WHERE LTRIM(RTRIM(call_id)) = :call_id)
      BEGIN
        UPDATE dbo.call_webhook_dtl
        SET
          campaign_id   = COALESCE(:campaign_id, campaign_id),
          direction     = COALESCE(:direction, direction),
          phone_number  = COALESCE(:phone_number, phone_number),
          status        = COALESCE(:status, status),
          triggered_at  = COALESCE(TRY_CONVERT(datetime2, :triggered_at), triggered_at),
          start_time    = COALESCE(TRY_CONVERT(datetime2, :start_time), start_time),
          end_time      = COALESCE(TRY_CONVERT(datetime2, :end_time), end_time),
          duration      = COALESCE(:duration, duration),
          transcript    = COALESCE(:transcript, transcript),
          summary       = COALESCE(:summary, summary),
          callee_name   = COALESCE(:callee_name, callee_name),
          category      = COALESCE(:category, category),
          ring_time     = COALESCE(:ring_time, ring_time),
          recording_url = COALESCE(:recording_url, recording_url)
        WHERE LTRIM(RTRIM(call_id)) = :call_id
      END
      ELSE
      BEGIN
        INSERT INTO dbo.call_webhook_dtl
          (call_id, campaign_id, direction, phone_number, status,
           triggered_at, start_time, end_time, duration,
           transcript, summary, created_at, callee_name, category, ring_time, recording_url)
        VALUES
          (:call_id, :campaign_id, :direction, :phone_number, :status,
           COALESCE(TRY_CONVERT(datetime2, :triggered_at), GETDATE()),
           TRY_CONVERT(datetime2, :start_time),
           TRY_CONVERT(datetime2, :end_time),
           :duration,
           :transcript, :summary, GETDATE(), :callee_name, :category, :ring_time, :recording_url)
      END
    `;

    await sequelize.query(upsertSql, {
      replacements: {
        call_id: cid,
        campaign_id,
        direction,
        phone_number,
        status,
        triggered_at: triggered_at ?? null,
        start_time: start_time ?? null,
        end_time: end_time ?? null,
        duration: duration != null && !isNaN(Number(duration)) ? Number(duration) : null,
        transcript: transcript ?? null,
        summary: summary ?? null,
        callee_name: callee_name ?? null,
        category: category ?? null,
        ring_time: ring_time != null && !isNaN(Number(ring_time)) ? Number(ring_time) : null,
        recording_url: recording_url ?? null,
      },
      type: QueryTypes.RAW,
    });
  } catch (err) {
    console.error("[CALLMATICS] saveCallWebhookDetails error for callId:", callId, err?.message);
  } finally {
    currentLockResolver();
    if (webhookSaveLocks.get(cid) === currentLock) {
      webhookSaveLocks.delete(cid);
    }
  }
};
exports.saveCallWebhookDetails = saveCallWebhookDetails;

// {CALLMATICS}
// exports.makeInsuranceRenewalCall = async function (req, res) {
//   let sequelize;
//   try {
//     console.log("===== RENEWAL NEW CODE v2 (ONLY VEHICLE NUMBER) =====");
//     sequelize = await dbname(req, req.headers.compcode);

//     const { vehicle_number } = req.body || {};

//     const cleanVeh = vehicle_number ? cleanVehicleNo(vehicle_number) : null;

//     if (!cleanVeh) {
//       return res.status(400).send({
//         success: false,
//         message: "vehicle_number is required",
//       });
//     }

//     const dbInfo = await sequelize.query("SELECT DB_NAME() AS db", {
//       type: QueryTypes.SELECT,
//     });
//     console.log(
//       "[RENEWAL] DB:",
//       dbInfo?.[0]?.db,
//       "compcode:",
//       req.headers.compcode,
//     );

//     const sql = `
//       SELECT TOP 1
//         r.UTD,
//         r.CUST_NAME,
//         r.CUST_MOB_NO,
//         r.POLICY_NAME,
//         CAST(r.POLICY_NUMBER AS varchar(100)) AS POLICY_NUMBER,
//         r.MODEL_NAME,
//         r.DSC_MOB_NO,
//         CONVERT(varchar(10), r.POLICY_START_DATE, 23) AS POLICY_START_DATE,
//         CONVERT(varchar(10), r.POLICY_END_DATE, 23)   AS POLICY_END_DATE,
//         r.TRAN_ID,

//         mst.UTD           AS MST_UTD,
//         mst.VEHICAL_REG_NO,
//         mst.LOC_CODE,
//         mst.EXPORT_TYPE
//       FROM dbo.INSU_RENEWAL r
//       INNER JOIN dbo.INSU_RENEWAL_MST mst
//         ON mst.UTD = r.TRAN_ID
//       WHERE ISNULL(mst.EXPORT_TYPE, 0) = 1
//         AND REPLACE(REPLACE(REPLACE(
//               UPPER(LTRIM(RTRIM(ISNULL(mst.VEHICAL_REG_NO,'')))),
//             ' ', ''), '-', ''), '/', '') = :cleanVeh
//       ORDER BY r.UTD DESC
//     `;

//     const rows = await sequelize.query(sql, {
//       replacements: { cleanVeh },
//       type: QueryTypes.SELECT,
//     });

//     console.log("[RENEWAL] Query result count:", rows?.length);
//     console.log("[RENEWAL] First row:", rows?.[0]);

//     if (!rows.length) {
//       return res.status(404).send({
//         success: false,
//         message: "Insurance renewal record not found for this vehicle_number",
//         debug: { cleanVeh, db: dbInfo?.[0]?.db },
//       });
//     }

//     const row = rows[0];

//     const phoneNumber = normalizePhone(row.CUST_MOB_NO);
//     if (!phoneNumber) {
//       return res.status(400).send({
//         success: false,
//         message: "Customer mobile not found",
//         debug: { UTD: row.UTD, CUST_NAME: row.CUST_NAME },
//       });
//     }

//     const locCode = row.LOC_CODE != null ? Number(row.LOC_CODE) : null;

//     const timeToAmPm = (hhmmss) => {
//       if (!hhmmss) return null;
//       const s = String(hhmmss).trim();
//       const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
//       if (!m) return s;
//       let h = parseInt(m[1], 10);
//       const min = m[2];
//       const ampm = h >= 12 ? "PM" : "AM";
//       h = h % 12;
//       if (h === 0) h = 12;
//       return `${h}:${min} ${ampm}`;
//     };

//     const toMDY = (d = new Date()) => {
//       const mm = d.getMonth() + 1;
//       const dd = d.getDate();
//       const yyyy = d.getFullYear();
//       return `${mm}/${dd}/${yyyy}`;
//     };

//     const cfgSQL = `
//       SELECT TOP 1
//         C.UTD,
//         C.INSU_COMPANY_NAME,
//         C.SALES_EXECUTIVE_NO,
//         C.SLOT1,
//         C.SLOT2,
//         C.SLOT3,
//         C.CAMPAIGN_ID,
//         CONVERT(varchar(8), C.CALLBACK_TIME) AS CALLBACK_TIME,
//         C.LOC_CODE,
//         C.STATUS
//       FROM dbo.INSU_CALLING_CONFIG C
//       WHERE (C.LOC_CODE = :locCode OR C.LOC_CODE IS NULL)
//       ORDER BY
//         CASE WHEN C.LOC_CODE = :locCode THEN 0 ELSE 1 END,
//         C.UTD DESC
//     `;

//     const cfgRows = await sequelize.query(cfgSQL, {
//       replacements: { locCode },
//       type: QueryTypes.SELECT,
//     });

//     const cfg = cfgRows?.[0] || null;

//     // =========================================================
//     // ✅ vehicle_model: use MODEL_NAME if available,
//     //    else use POLICY_NAME, else use vehicle_number as fallback
//     // =========================================================
//     const resolvedVehicleModel = (() => {
//       if (row.MODEL_NAME != null && String(row.MODEL_NAME).trim() !== "") {
//         return String(row.MODEL_NAME).trim();
//       }
//       // fallback 1: try POLICY_NAME
//       if (row.POLICY_NAME != null && String(row.POLICY_NAME).trim() !== "") {
//         return String(row.POLICY_NAME).trim();
//       }
//       // fallback 2: use vehicle number itself
//       if (
//         row.VEHICAL_REG_NO != null &&
//         String(row.VEHICAL_REG_NO).trim() !== ""
//       ) {
//         return String(row.VEHICAL_REG_NO).trim();
//       }
//       return null;
//     })();

//     // =========================================================
//     // ✅ transferNumber: First check DSC_NO from INSU_RENEWAL,
//     //    then fallback to SALES_EXECUTIVE_NO from config
//     // =========================================================
//     const resolvedTransferNumber = (() => {
//       if (row.DSC_MOB_NO != null && String(row.DSC_NO).trim() !== "") {
//         return String(row.DSC_MOB_NO).trim();
//       }
//       // fallback: try config SALES_EXECUTIVE_NO
//       if (cfg?.SALES_EXECUTIVE_NO != null) {
//         return String(cfg.SALES_EXECUTIVE_NO);
//       }
//       return null;
//     })();

//     const variables = {
//       showroom_name: row.POLICY_NAME ?? cfg?.INSU_COMPANY_NAME ?? null,
//       // ✅ Now uses fallback logic instead of null
//       vehicle_model: resolvedVehicleModel,
//       vehicle_number:
//         row.VEHICAL_REG_NO != null && String(row.VEHICAL_REG_NO).trim() !== ""
//           ? String(row.VEHICAL_REG_NO).trim()
//           : cleanVeh, // fallback to input if DB empty
//       insurance_expiry_date: formatDate(row.POLICY_END_DATE) || null,
//       transferNumber: resolvedTransferNumber,
//       callback_date: toMDY(new Date()),
//       callback_time: timeToAmPm(cfg?.CALLBACK_TIME) ?? null,
//       callee_name:
//         row.CUST_NAME != null && String(row.CUST_NAME).trim() !== ""
//           ? String(row.CUST_NAME).trim()
//           : null,
//     };

//     const finalCampaignId = cfg?.CAMPAIGN_ID || null;

//     // =========================================================
//     // ✅ vehicle_model is NOT in missing check (has fallback)
//     // ✅ vehicle_number is NOT in missing check (has fallback)
//     // =========================================================
//     const missing = [];
//     if (!variables.showroom_name)
//       missing.push(
//         "showroom_name (POLICY_NAME/INSU_CALLING_CONFIG.INSU_COMPANY_NAME)",
//       );
//     if (!variables.insurance_expiry_date)
//       missing.push("insurance_expiry_date (POLICY_END_DATE)");
//     if (!variables.transferNumber)
//       missing.push(
//         "transferNumber (INSU_RENEWAL.DSC_NO/INSU_CALLING_CONFIG.SALES_EXECUTIVE_NO)",
//       );
//     if (!variables.callback_time)
//       missing.push("callback_time (INSU_CALLING_CONFIG.CALLBACK_TIME)");
//     if (!variables.callee_name) missing.push("callee_name (CUST_NAME)");
//     if (!finalCampaignId)
//       missing.push("campaign_id (INSU_CALLING_CONFIG.CAMPAIGN_ID)");

//     if (missing.length) {
//       return res.status(400).send({
//         success: false,
//         message:
//           "Dynamic values missing. Please configure INSU_CALLING_CONFIG for this LOC_CODE (or global LOC_CODE NULL).",
//         missingFields: missing,
//         debug: {
//           vehicle_number,
//           cleanVeh,
//           locCode,
//           foundConfig: !!cfg,
//           cfgUTD: cfg?.UTD || null,
//           // ✅ Show what model was resolved to help debug
//           resolvedVehicleModel,
//           rawModelName: row.MODEL_NAME,
//           // ✅ Show what transfer number was resolved
//           resolvedTransferNumber,
//           rawDscNo: row.DSC_NO,
//           rawSalesExecNo: cfg?.SALES_EXECUTIVE_NO || null,
//         },
//       });
//     }

//     console.log("[RENEWAL] Variables prepared:", {
//       ...variables,
//       // mask phone for log
//       transferNumber: variables.transferNumber
//         ? "****" + String(variables.transferNumber).slice(-4)
//         : null,
//     });

//     // Call trigger
//     const callResult = await triggerSingleCall(
//       phoneNumber,
//       variables,
//       finalCampaignId,
//     );

//     const callId = callResult?.callId || callResult?.calls?.[0]?.callId;
//     const calledPhone =
//       callResult?.phoneNumber ||
//       callResult?.calls?.[0]?.phoneNumber ||
//       phoneNumber;

//     // Save call log
//     if (callId) {
//       try {
//         // 1) Save to call_Id_dtl
//         await sequelize.query(
//           `INSERT INTO dbo.call_Id_dtl (mob_no, call_id, call_type)
//            VALUES (:mob_no, :call_id, 'INSURANCE_RENEWAL')`,
//           {
//             replacements: { mob_no: calledPhone, call_id: callId },
//             type: QueryTypes.INSERT,
//           },
//         );

//         // 2) Initial save to call_webhook_dtl
//         await saveCallWebhookDetails(sequelize, callId, {
//           campaignId: finalCampaignId,
//           phoneNumber: calledPhone,
//           calleeName: variables.callee_name || null,
//           status: 'INITIATED',
//           direction: 'outbound',
//           triggeredAt: new Date(),
//           startTime: new Date(),
//           category: 'INSURANCE_RENEWAL',
//         });

//         // 3) Save to FOLLOWUP_DETAILS so it links with calling history
//         const mstUtd = row.MST_UTD || row.TRAN_ID;
//         const expType = row.EXPORT_TYPE != null ? Number(row.EXPORT_TYPE) : 1;

//         if (mstUtd) {
//           await sequelize.query(
//             `INSERT INTO dbo.FOLLOWUP_DETAILS (
//               TRAN_ID,
//               EXPORT_TYPE,
//               FOLLOWUP_STATUS,
//               FOLLOWUP_DATE,
//               FOLLOWUP_TIME,
//               LAST_FOLLOWUP_DATE,
//               REMARKS,
//               CALL_ID
//              ) VALUES (
//               :tranId,
//               :exportType,
//               'AI_CALL_INITIATED',
//               CAST(GETDATE() AS date),
//               CAST(GETDATE() AS time),
//               CAST(GETDATE() AS date),
//               'AI Call triggered manually',
//               :callId
//              )`,
//             {
//               replacements: {
//                 tranId: mstUtd,
//                 exportType: expType,
//                 callId: String(callId),
//               },
//               type: QueryTypes.INSERT,
//             },
//           );
//           console.log("[RENEWAL] Saved callId to FOLLOWUP_DETAILS:", callId);
//         }
//       } catch (logErr) {
//         console.error("[RENEWAL] Call log or followup save failed:", logErr?.message);
//       }
//     }

//     // =========================================================
//     // ✅ AUTO WHATSAPP: Call trigger ke baad wait kro, call complete
//     //    ho jae tb hi WhatsApp message bhejo
//     // =========================================================
//     setImmediate(async () => {
//       let bgSeq;
//       try {
//         console.log("[RENEWAL] WhatsApp & status sync process started in background");
//         bgSeq = await dbname(req, req.headers.compcode);

//         // ✅ Call ke status check kro - complete/ended/finished ho gaya ya nahi
//         const checkCallStatus = async (cId, maxAttempts = 60) => {
//           let attempts = 0;

//           while (attempts < maxAttempts) {
//             try {
//               // API se call status check kro
//               const statusResult = await getCallStatus1(cId);

//               console.log(`[RENEWAL] Call Status Check (Attempt ${attempts + 1}/${maxAttempts}):`, {
//                 callId: cId,
//                 status: statusResult?.status,
//                 duration: statusResult?.duration,
//               });

//               if (statusResult && bgSeq) {
//                 await saveCallWebhookDetails(bgSeq, cId, {
//                   ...statusResult,
//                   phoneNumber: calledPhone,
//                   calleeName: variables.callee_name || null,
//                   campaignId: finalCampaignId,
//                 });
//               }

//               const st = String(statusResult?.status || '').toUpperCase();

//               // ✅ Call complete/ended/finished ho gaya to return kro
//               if (['COMPLETED', 'ENDED', 'FINISHED', 'FAILED', 'BUSY', 'NO_ANSWER', 'CANCELED'].includes(st)) {
//                 console.log("[RENEWAL] ✅ Call Status:", statusResult?.status, "- Checking summary & Proceeding");

//                 // Agar call complete hui par summary nahi aayi, 1-2 brief retries karo taaki Callmatic summary process kar sake
//                 if (['COMPLETED', 'ENDED', 'FINISHED'].includes(st) && !statusResult?.summary) {
//                   for (let sAttempt = 1; sAttempt <= 3; sAttempt++) {
//                     await new Promise(resolve => setTimeout(resolve, 3000));
//                     try {
//                       const retryRes = await getCallStatus1(cId);
//                       if (retryRes && bgSeq) {
//                         await saveCallWebhookDetails(bgSeq, cId, {
//                           ...retryRes,
//                           phoneNumber: calledPhone,
//                           calleeName: variables.callee_name || null,
//                           campaignId: finalCampaignId,
//                         });
//                       }
//                       if (retryRes?.summary || retryRes?.analysis?.summary || retryRes?.call_analysis?.summary) {
//                         console.log(`[RENEWAL] ✅ Summary received on retry attempt ${sAttempt}`);
//                         break;
//                       }
//                     } catch (e) {
//                       // ignore retry error
//                     }
//                   }
//                 }

//                 return {
//                   success: true,
//                   status: statusResult?.status,
//                   duration: statusResult?.duration,
//                 };
//               }

//               // Agar call still ringing/active hai to wait kro
//               if (['RINGING', 'ACTIVE', 'IN_PROGRESS', 'INITIATED', 'QUEUED'].includes(st)) {
//                 console.log(`[RENEWAL] Call still ${statusResult?.status}... waiting...`);
//               }

//               // 2 second wait kro phir se check kro
//               await new Promise(resolve => setTimeout(resolve, 2000));
//               attempts++;
//             } catch (checkErr) {
//               console.error(`[RENEWAL] Status check error (Attempt ${attempts + 1}):`, checkErr?.message);
//               attempts++;
//               // Error aaye to 2 second wait kro
//               await new Promise(resolve => setTimeout(resolve, 2000));
//             }
//           }

//           // Max attempts ho gaye
//           console.log("[RENEWAL] ⚠️ Max status check attempts reached (120 seconds timeout)");
//           return {
//             success: false,
//             status: 'TIMEOUT',
//             message: 'Call status check timeout - WhatsApp abhi bhejenge'
//           };
//         };

//         // ✅ WAIT FOR CALL TO COMPLETE
//         console.log("[RENEWAL] Waiting for call to complete...");
//         if (callId) {
//           const callStatusResult = await checkCallStatus(callId);

//           console.log("[RENEWAL] Call Status Result:", {
//             success: callStatusResult?.success,
//             status: callStatusResult?.status,
//             message: callStatusResult?.message,
//           });

//           // ✅ Agar call complete nahi hua to wait kro additional 5 seconds
//           if (!callStatusResult?.success) {
//             console.log("[RENEWAL] Waiting additional 5 seconds before sending WhatsApp...");
//             await new Promise(resolve => setTimeout(resolve, 5000));
//           }
//         }

//         // ✅ NOW SEND WHATSAPP (jab call complete ho gaya)
//         console.log("[RENEWAL] 🔔 Calling SendInsuranceRenewalWhatsAppToCustomer...");

//         const compcode = req.headers.compcode;
//         const mstUtd = row.MST_UTD;

//         const reqObj = {
//           headers: { compcode },
//           body: { tranIds: [Number(mstUtd)] },
//         };

//         let waResult;
//         const resObj = {
//           status: (code) => ({
//             send: (data) => {
//               waResult = { statusCode: code, data };
//               return waResult;
//             },
//             json: (data) => {
//               waResult = { statusCode: code, data };
//               return waResult;
//             },
//           }),
//         };

//         await exports.SendInsuranceRenewalWhatsAppToCustomer(reqObj, resObj);

//         console.log("[RENEWAL] ✅ Auto-WhatsApp completed via SendInsuranceRenewalWhatsAppToCustomer:", {
//           compcode,
//           mstUtd,
//           callId: callId || null,
//           waResult,
//           timestamp: new Date().toISOString(),
//         });

//       } catch (waErr) {
//         console.error("[RENEWAL] ❌ Auto-WhatsApp failed:", {
//           error: waErr?.message || waErr,
//           code: waErr?.code,
//           timestamp: new Date().toISOString(),
//         });
//       } finally {
//         if (bgSeq) {
//           try {
//             await bgSeq.close();
//           } catch (_) { }
//         }
//       }
//     });

//     return res.status(200).send({
//       success: true,
//       message: "Insurance Renewal AI Call Triggered Successfully",
//       data: callResult,
//       variables,
//       renewal: {
//         UTD: row.UTD,
//         TRAN_ID: row.TRAN_ID,
//         MST_UTD: row.MST_UTD,
//         LOC_CODE: row.LOC_CODE,
//         CUST_NAME: row.CUST_NAME,
//         VEHICAL_REG_NO: row.VEHICAL_REG_NO,
//         POLICY_NUMBER: row.POLICY_NUMBER,
//         POLICY_END_DATE: row.POLICY_END_DATE,
//         callId: callId || null,
//         // ✅ Show model source for transparency
//         MODEL_NAME_SOURCE: row.MODEL_NAME ? "DB" : "FALLBACK",
//         // ✅ Show transfer number source for transparency
//         TRANSFER_NUMBER_SOURCE: row.DSC_NO ? "DSC_NO" : "SALES_EXECUTIVE_NO",
//       },
//     });
//   } catch (err) {
//     console.error("[RENEWAL] ERROR:", err);
//     return res.status(500).send({
//       success: false,
//       message: err?.message || "Server error",
//     });
//   } finally {
//     if (sequelize) {
//       try {
//         await sequelize.close();
//       } catch (_) { }
//     }
//   }
// };
// {BONVOICE}
exports.makeInsuranceRenewalCall = async function (req, res) {
  let sequelize;
  try {
    console.log("===== RENEWAL NEW CODE v2 (ONLY VEHICLE NUMBER) =====");
    sequelize = await dbname(req, req.headers.compcode);

    const { vehicle_number } = req.body || {};
    const cleanVeh = vehicle_number ? cleanVehicleNo(vehicle_number) : null;

    if (!cleanVeh) {
      return res.status(400).send({
        success: false,
        message: "vehicle_number is required",
      });
    }

    const dbInfo = await sequelize.query("SELECT DB_NAME() AS db", {
      type: QueryTypes.SELECT,
    });
    console.log(
      "[RENEWAL] DB:",
      dbInfo?.[0]?.db,
      "compcode:",
      req.headers.compcode
    );

    const sql = `
      SELECT TOP 1
        r.UTD,
        r.CUST_NAME,
        r.CUST_MOB_NO,
        r.POLICY_NAME,
        CAST(r.POLICY_NUMBER AS varchar(100)) AS POLICY_NUMBER,
        r.MODEL_NAME,
        CONVERT(varchar(10), r.POLICY_START_DATE, 23) AS POLICY_START_DATE,
        CONVERT(varchar(10), r.POLICY_END_DATE, 23)   AS POLICY_END_DATE,
        r.TRAN_ID,

        mst.UTD           AS MST_UTD,
        mst.VEHICAL_REG_NO,
        mst.LOC_CODE,
        mst.EXPORT_TYPE
      FROM dbo.INSU_RENEWAL r
      INNER JOIN dbo.INSU_RENEWAL_MST mst
        ON mst.UTD = r.TRAN_ID
      WHERE ISNULL(mst.EXPORT_TYPE, 0) = 1
        AND REPLACE(REPLACE(REPLACE(
              UPPER(LTRIM(RTRIM(ISNULL(mst.VEHICAL_REG_NO,'')))),
            ' ', ''), '-', ''), '/', '') = :cleanVeh
      ORDER BY r.UTD DESC
    `;

    const rows = await sequelize.query(sql, {
      replacements: { cleanVeh },
      type: QueryTypes.SELECT,
    });

    console.log("[RENEWAL] Query result count:", rows?.length);
    console.log("[RENEWAL] First row:", rows?.[0]);

    if (!rows.length) {
      return res.status(404).send({
        success: false,
        message: "Insurance renewal record not found for this vehicle_number",
        debug: { cleanVeh, db: dbInfo?.[0]?.db },
      });
    }

    const row = rows[0];

    const phoneNumber = normalizePhone(row.CUST_MOB_NO);
    if (!phoneNumber) {
      return res.status(400).send({
        success: false,
        message: "Customer mobile not found",
        debug: { UTD: row.UTD, CUST_NAME: row.CUST_NAME },
      });
    }

    const locCode = row.LOC_CODE != null ? Number(row.LOC_CODE) : null;

    const timeToAmPm = (hhmmss) => {
      if (!hhmmss) return null;
      const s = String(hhmmss).trim();
      const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
      if (!m) return s;
      let h = parseInt(m[1], 10);
      const min = m[2];
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

    const cfgSQL =
      locCode == null
        ? `
      SELECT TOP 1
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
      WHERE C.LOC_CODE IS NULL
      ORDER BY C.UTD DESC
    `
        : `
      SELECT TOP 1
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
        C.UTD DESC
    `;

    const cfgRows = await sequelize.query(cfgSQL, {
      replacements: { locCode },
      type: QueryTypes.SELECT,
    });

    const cfg = cfgRows?.[0] || null;

    const resolvedVehicleModel = (() => {
      if (row.MODEL_NAME != null && String(row.MODEL_NAME).trim() !== "") {
        return String(row.MODEL_NAME).trim();
      }
      if (row.POLICY_NAME != null && String(row.POLICY_NAME).trim() !== "") {
        return String(row.POLICY_NAME).trim();
      }
      if (
        row.VEHICAL_REG_NO != null &&
        String(row.VEHICAL_REG_NO).trim() !== ""
      ) {
        return String(row.VEHICAL_REG_NO).trim();
      }
      return null;
    })();

    const variables = {
      showroom_name: cfg?.INSU_COMPANY_NAME ?? row.POLICY_NAME ?? null,
      vehicle_model: resolvedVehicleModel,
      vehicle_number:
        row.VEHICAL_REG_NO != null && String(row.VEHICAL_REG_NO).trim() !== ""
          ? String(row.VEHICAL_REG_NO).trim()
          : cleanVeh,
      insurance_expiry_date: formatDate(row.POLICY_END_DATE) || null,
      transferNumber:
        cfg?.SALES_EXECUTIVE_NO != null ? String(cfg.SALES_EXECUTIVE_NO) : null,
      callback_date: toMDY(new Date()),
      callback_time: timeToAmPm(cfg?.CALLBACK_TIME) ?? null,
      callee_name:
        row.CUST_NAME != null && String(row.CUST_NAME).trim() !== ""
          ? String(row.CUST_NAME).trim()
          : null,
    };

    const promptNameRaw = cfg?.CAMPAIGN_ID || null;
    const availablePrompts = ["default", "service", "Insurance"];

    const bonvoicePromptName = availablePrompts.includes(promptNameRaw)
      ? promptNameRaw
      : "Insurance";

    const missing = [];
    if (!variables.showroom_name)
      missing.push(
        "showroom_name (POLICY_NAME/INSU_CALLING_CONFIG.INSU_COMPANY_NAME)"
      );
    if (!variables.insurance_expiry_date)
      missing.push("insurance_expiry_date (POLICY_END_DATE)");
    if (!variables.transferNumber)
      missing.push("transferNumber (INSU_CALLING_CONFIG.SALES_EXECUTIVE_NO)");
    if (!variables.callback_time)
      missing.push("callback_time (INSU_CALLING_CONFIG.CALLBACK_TIME)");
    if (!variables.callee_name) missing.push("callee_name (CUST_NAME)");
    if (!bonvoicePromptName)
      missing.push("promptName (INSU_CALLING_CONFIG.CAMPAIGN_ID)");

    if (missing.length) {
      return res.status(400).send({
        success: false,
        message:
          "Dynamic values missing. Please configure INSU_CALLING_CONFIG for this LOC_CODE (or global LOC_CODE NULL).",
        missingFields: missing,
        debug: {
          vehicle_number,
          cleanVeh,
          locCode,
          foundConfig: !!cfg,
          cfgUTD: cfg?.UTD || null,
          resolvedVehicleModel,
          rawModelName: row.MODEL_NAME,
          promptNameRaw,
          bonvoicePromptName,
        },
      });
    }

    console.log("[RENEWAL] Variables prepared:", {
      ...variables,
      transferNumber: variables.transferNumber
        ? "****" + String(variables.transferNumber).slice(-4)
        : null,
    });

    const bonvoiceProgram = "INSURANCE_RENEWAL";

    // ✅ CALL BONVOICE (UPDATED: 2-step so callId mil jaye)
    const leadPayload = {
      name: variables.callee_name,
      phone: phoneNumber,
      email: null,
      company: variables.showroom_name,
      program: bonvoiceProgram,

      column1: String(row.MST_UTD),
      column2: variables.vehicle_number,
      column3: variables.insurance_expiry_date,
      column7: variables.transferNumber
    };

    const callOptions = {
      promptName: bonvoicePromptName,
      program: bonvoiceProgram,
    };

    // NOTE: createLeadThenTriggerCall ko aapke bonvoice helper se import hona chahiye
    // const { createLeadThenTriggerCall } = require("...");

    const callResult = await createLeadThenTriggerCall(leadPayload, callOptions);

    const bonvoiceLeadId = callResult?.leadId || null;
    const callId = callResult?.callId || null;
    const calledPhone = phoneNumber;

    console.log("[RENEWAL] Bonvoice call triggered:", {
      leadId: bonvoiceLeadId,
      callId: callId,
      phone: calledPhone,
    });

    // ✅ Save leadId in FOLLOWUP_DETAILS for webhook reference
    if (bonvoiceLeadId) {
      try {
        await sequelize.query(
          `INSERT INTO dbo.FOLLOWUP_DETAILS
           (TRAN_ID, bonvoice_lead_id, FOLLOWUP_STATUS, FOLLOWUP_DATE, CREATED_AT)
           VALUES (:tranId, :leadId, 'AI_CALL_INITIATED', GETDATE(), GETDATE())`,
          {
            replacements: {
              tranId: row.MST_UTD,
              leadId: bonvoiceLeadId,
            },
            type: QueryTypes.INSERT,
          }
        );
        console.log("[RENEWAL] Bonvoice leadId saved:", bonvoiceLeadId);
      } catch (fuErr) {
        console.warn("[RENEWAL] leadId save failed:", fuErr?.message);
      }
    }

    // ✅ NEW: Save callId ↔ leadId mapping for webhook enrichment
    if (callId) {
      try {
        const mergeMapSql = `
          MERGE dbo.call_webhook_dtl AS T
          USING (SELECT :call_id AS call_id) AS S
          ON T.call_id = S.call_id
          WHEN MATCHED THEN
            UPDATE SET
              lead_id      = ISNULL(:lead_id, T.lead_id),
              phone_number = ISNULL(:phone_number, T.phone_number),
              callee_name  = ISNULL(:callee_name, T.callee_name),
              campaign_id  = ISNULL(:campaign_id, T.campaign_id)
          WHEN NOT MATCHED THEN
            INSERT (call_id, lead_id, phone_number, callee_name, campaign_id, created_at)
            VALUES (:call_id, :lead_id, :phone_number, :callee_name, :campaign_id, GETDATE());
        `;

        await sequelize.query(mergeMapSql, {
          replacements: {
            call_id: String(callId).trim(),
            lead_id: bonvoiceLeadId ? String(bonvoiceLeadId).trim() : null,
            phone_number: String(calledPhone).substring(0, 15),
            callee_name: variables.callee_name
              ? String(variables.callee_name).substring(0, 100)
              : null,
            campaign_id: bonvoicePromptName
              ? String(bonvoicePromptName).substring(0, 100)
              : null,
          },
          type: QueryTypes.RAW,
        });

        console.log("[RENEWAL] call_webhook_dtl mapping saved:", {
          callId,
          bonvoiceLeadId,
        });
      } catch (mapErr) {
        console.warn(
          "[RENEWAL] call_webhook_dtl mapping save failed:",
          mapErr?.message
        );
      }

      // existing call log
      try {
        await sequelize.query(
          `INSERT INTO dbo.call_Id_dtl (mob_no, call_id, call_type)
           VALUES (:mob_no, :call_id, 'INSURANCE_RENEWAL')`,
          {
            replacements: { mob_no: calledPhone, call_id: callId },
            type: QueryTypes.INSERT,
          }
        );
      } catch (logErr) {
        console.error("[RENEWAL] Call log save failed:", logErr?.message);
      }
    }

    return res.status(200).send({
      success: true,
      message: "Insurance Renewal AI Call Triggered Successfully",
      data: callResult,
      variables,
      bonvoice: {
        leadId: bonvoiceLeadId,
        callId: callId,
        callStatus: callResult?.call?.callStatus || callResult?.callStatus || null,
        promptName: bonvoicePromptName,
        program: bonvoiceProgram,
      },
      renewal: {
        UTD: row.UTD,
        TRAN_ID: row.TRAN_ID,
        MST_UTD: row.MST_UTD,
        LOC_CODE: row.LOC_CODE,
        CUST_NAME: row.CUST_NAME,
        VEHICAL_REG_NO: row.VEHICAL_REG_NO,
        POLICY_NUMBER: row.POLICY_NUMBER,
        POLICY_END_DATE: row.POLICY_END_DATE,
        callId: callId || null,
        MODEL_NAME_SOURCE: row.MODEL_NAME ? "DB" : "FALLBACK",
      },
    });
  } catch (err) {
    console.error("[RENEWAL] ERROR:", err);
    return res.status(500).send({
      success: false,
      message: err?.message || "Server error",
      debug: err?.data || null,
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) {}
    }
  }
};

exports.callmaticWebhook = async function (req, res) {
  let sequelize;

  try {
    // ---- helpers ----
    const pick = (...vals) =>
      vals.find((v) => v !== undefined && v !== null && v !== "");

    const extractCallId = (b = {}) =>
      pick(
        b.callId,
        b.call_id,
        b.data?.callId,
        b.data?.call_id,
        b.payload?.callId,
        b.payload?.call_id,
      ) || null;

    const cleanVehicleNo = (v) =>
      String(v || "")
        .toUpperCase()
        .replace(/\s+/g, "")
        .replace(/[-/]/g, "")
        .trim();

    // const toJsonString = (v) => {
    //   if (v === undefined || v === null) return null;
    //   if (typeof v === "string") return v;
    //   try {
    //     return JSON.stringify(v);
    //   } catch {
    //     return String(v);
    //   }
    // };

    // ✅ SQL Server safe datetime string: YYYY-MM-DD HH:mm:ss.SSS
    // const toSqlDateTimeOrNull = (v) => {
    //   if (!v) return null;
    //   const d = typeof v === "number" ? new Date(v * 1000) : new Date(v);
    //   if (isNaN(d.getTime())) return null;

    //   const pad2 = (n) => String(n).padStart(2, "0");
    //   const pad3 = (n) => String(n).padStart(3, "0");

    //   return (
    //     `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ` +
    //     `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}.` +
    //     `${pad3(d.getMilliseconds())}`
    //   );
    // };

    // ---- compcode resolve (header / params / query) ----
    const compcode = pick(
      req.headers.compcode,
      req.params.compcode,
      req.query.compcode,
    );
    if (!compcode) {
      return res.status(400).json({
        Status: false,
        Message:
          "compcode is required (send in header OR /:compcode param OR query)",
      });
    }

    sequelize = await dbname(req, compcode);

    const body = req.body || {};
    const callId = extractCallId(body);

    // ============================================================
    // ✅ MODE-2 (FRONTEND): vehicle_number -> MST -> FOLLOWUP_DETAILS -> call_webhook_dtl
    // ============================================================
    const vehicle_number = pick(
      body.vehicle_number,
      body.VEHICAL_REG_NO,
      body.vehicleNumber,
    );

    // Agar callId nahi hai but vehicle_number hai => DB linking se details do
    if (!callId && vehicle_number) {
      const cleanVeh = cleanVehicleNo(vehicle_number);

      const normExpr = (col) =>
        `REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(${col},'')))),' ',''),'-',''),'/','')`;

      // 1) Find MST by vehicle reg no
      const mstSql = `
            SELECT TOP 1
              mst.UTD,
              mst.VEHICAL_REG_NO
            FROM dbo.INSU_RENEWAL_MST mst
            WHERE ISNULL(mst.EXPORT_TYPE,0) = 1
              AND ${normExpr("mst.VEHICAL_REG_NO")} = :cleanVeh
            ORDER BY mst.UTD DESC
          `;

      const mstRows = await sequelize.query(mstSql, {
        replacements: { cleanVeh },
        type: QueryTypes.SELECT,
      });

      if (!mstRows?.length) {
        return res.status(404).json({
          Status: false,
          Message: "Vehicle not found in INSU_RENEWAL_MST",
          vehicle_number,
        });
      }

      const mst = mstRows[0];

      // 2) Find all call_ids for this vehicle
      const allCallIdsSql = `
            SELECT DISTINCT LTRIM(RTRIM(x.call_id)) AS call_id
            FROM (
              -- 1) Any FOLLOWUP_DETAILS for all MST UTDs of this vehicle
              SELECT fd.call_id
              FROM dbo.FOLLOWUP_DETAILS fd
              INNER JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = fd.TRAN_ID
              WHERE fd.call_id IS NOT NULL AND LTRIM(RTRIM(fd.call_id)) <> ''
                AND REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(m.VEHICAL_REG_NO,'')))),' ',''),'-',''),'/','') = :cleanVeh

              UNION

              -- 2) Any FOLLOWUP_DETAILS for INSU_RENEWAL of this vehicle
              SELECT fd.call_id
              FROM dbo.FOLLOWUP_DETAILS fd
              INNER JOIN dbo.INSU_RENEWAL r ON (r.TRAN_ID = fd.TRAN_ID OR r.UTD = fd.TRAN_ID)
              WHERE fd.call_id IS NOT NULL AND LTRIM(RTRIM(fd.call_id)) <> ''
                AND REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(r.VEHICAL_REG_NO,'')))),' ',''),'-',''),'/','') = :cleanVeh
            ) x
            WHERE x.call_id IS NOT NULL AND LTRIM(RTRIM(x.call_id)) <> ''
          `;

      const callIdRows = await sequelize.query(allCallIdsSql, {
        replacements: { cleanVeh },
        type: QueryTypes.SELECT,
      });

      const callIds = (callIdRows || []).map((r) => r.call_id).filter(Boolean);

      if (!callIds.length) {
        return res.status(200).json({
          Status: true,
          Message: "No call records found for this vehicle",
          vehicle_number,
          mst: {
            UTD: mst.UTD,
            VEHICAL_REG_NO: mst.VEHICAL_REG_NO,
          },
          data: [],
        });
      }

      // 3) Auto-sync pending or incomplete call details from Callmatic API
      for (const cId of callIds) {
        try {
          const existing = await sequelize.query(
            `SELECT TOP 1 status, transcript, summary, duration, recording_url FROM dbo.call_webhook_dtl WHERE call_id = :callId`,
            { replacements: { callId: cId }, type: QueryTypes.SELECT }
          );
          const st = String(existing?.[0]?.status || '').toUpperCase();
          const isTerminal = ['COMPLETED', 'ENDED', 'FINISHED', 'FAILED', 'BUSY', 'NO_ANSWER', 'CANCELED'].includes(st);
          const isMissingSummary = ['COMPLETED', 'ENDED', 'FINISHED'].includes(st) && (!existing?.[0]?.summary || String(existing[0].summary).trim() === '');

          if (!existing?.length || !isTerminal || isMissingSummary) {
            const apiRes = await getCallStatus1(cId);
            if (apiRes) {
              await saveCallWebhookDetails(sequelize, cId, apiRes);
            }
          }
        } catch (syncErr) {
          console.error("[WEBHOOK SYNC] Error fetching status for callId:", cId, syncErr?.message);
        }
      }

      // 4) Fetch webhook details from call_webhook_dtl
      const webhookSql = `
            ;WITH RankedCalls AS (
              SELECT
                w.call_id,
                w.campaign_id,
                w.direction,
                w.phone_number,
                w.status,
                CONVERT(varchar(19), w.triggered_at, 120) AS triggered_at,
                CONVERT(varchar(19), w.start_time, 120)   AS start_time,
                CONVERT(varchar(19), w.end_time, 120)     AS end_time,
                w.duration,
                w.ring_time,
                w.summary,
                w.transcript,
                w.callee_name,
                w.category,
                w.recording_url,
                CONVERT(varchar(19), w.created_at, 120)   AS created_at,
                w.created_at                              AS raw_created_at,
                w.start_time                              AS raw_start_time,
                w.triggered_at                            AS raw_triggered_at,
                ROW_NUMBER() OVER (
                  PARTITION BY LTRIM(RTRIM(w.call_id))
                  ORDER BY 
                    CASE WHEN UPPER(ISNULL(w.status,'')) IN ('COMPLETED','ENDED','FINISHED') THEN 1 ELSE 2 END,
                    COALESCE(w.end_time, w.start_time, w.triggered_at, w.created_at) DESC,
                    w.created_at DESC
                ) AS rn
              FROM dbo.call_webhook_dtl w
              WHERE w.call_id IN (:callIds)
            )
            SELECT TOP 50
              call_id,
              campaign_id,
              direction,
              phone_number,
              status,
              triggered_at,
              start_time,
              end_time,
              duration,
              ring_time,
              summary,
              transcript,
              callee_name,
              category,
              recording_url,
              created_at
            FROM RankedCalls
            WHERE rn = 1
            ORDER BY COALESCE(raw_start_time, raw_triggered_at, raw_created_at) DESC, raw_created_at DESC
          `;
      const webhookRows = await sequelize.query(webhookSql, {
        replacements: { callIds },
        type: QueryTypes.SELECT,
      });

      const cleanedWebhookRows = (webhookRows || []).map((row) => ({
        ...row,
        summary: extractCleanSummary(row.summary),
      }));

      return res.status(200).json({
        Status: true,
        Message: "Webhook details fetched by vehicle_number",
        vehicle_number,
        mst: {
          UTD: mst.UTD,
          VEHICAL_REG_NO: mst.VEHICAL_REG_NO,
        },
        data: cleanedWebhookRows,
      });
    }

    // ============================================================
    // ✅ MODE-1 (WEBHOOK RECEIVER): callId required
    // ============================================================
    if (!callId) {
      return res.status(400).json({
        Status: false,
        Message:
          "callId is required (not found in webhook payload) OR send vehicle_number",
        ReceivedKeys: Object.keys(body || {}),
      });
    }

    // 1) Get latest call info from Callmatic API (fallback / enrichment)
    let callDetails = null;
    try {
      callDetails = await getCallStatus1(callId);
    } catch (e) {
      callDetails = null;
      console.error("[WEBHOOK] getCallStatus failed:", e?.message);
    }

    const payload = req.body || {};
    const api = callDetails || {};

    await saveCallWebhookDetails(sequelize, callId, { ...api, ...payload });

    const status = pick(payload.status, payload.call?.status, api.status, api.call?.status, api.callStatus) || null;

    // ============================================================
    // ✅ AUTO WHATSAPP: Call complete hone ke baad WhatsApp bhejo
    //    status === "completed" AND call type = INSURANCE_RENEWAL
    // ============================================================
    let whatsappTriggered = false;
    if (status && status.toLowerCase() === "completed" && phone_number) {
      setImmediate(async () => {
        let waSeq;
        try {
          waSeq = await dbname(req, compcode);

          const cleanPhone = String(phone_number).replace(/\D/g, "").slice(-10);
          const renewalSql = `
                SELECT TOP 1
                  m.UTD        AS MST_UTD,
                  m.VEHICAL_REG_NO,
                  m.LOC_CODE,
                  r.CUST_NAME,
                  r.CUST_MOB_NO,
                  r.POLICY_NAME,
                  r.MODEL_NAME
                FROM dbo.INSU_RENEWAL r
                INNER JOIN dbo.INSU_RENEWAL_MST m
                  ON m.UTD = r.TRAN_ID
                WHERE ISNULL(m.EXPORT_TYPE, 0) = 1
                  AND r.EXPORT_TYPE = 1
                  AND RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(r.CUST_MOB_NO,''),' ',''),'-',''),'+',''), 10) = :cleanPhone
                ORDER BY r.UTD DESC
              `;

          const renewalRows = await waSeq.query(renewalSql, {
            replacements: { cleanPhone },
            type: QueryTypes.SELECT,
          });

          if (!renewalRows?.length) {
            console.log("[AUTO-WA] No INSURANCE_RENEWAL record found for phone:", cleanPhone);
            return;
          }

          const wRow = renewalRows[0];
          const mstUtd = wRow.MST_UTD;

          let questionsCount = 0;
          try {
            const qRows = await waSeq.query(`
                  SELECT COUNT(*) AS cnt FROM dbo.INSU_QUESTIONS
                  WHERE (Is_Active = '1' OR Is_Active = 'true' OR Is_Active = 'Y' OR Is_Active IS NULL)
                `, { type: QueryTypes.SELECT });
            questionsCount = qRows?.[0]?.cnt ?? 0;
          } catch (_) { }

          const encodedCompCode = Buffer.from(compcode).toString("base64");
          const encodedUTD = Buffer.from(String(mstUtd)).toString("base64");
          const customerWebLink = `${process.env.BASE_URL || ""}/Crm/InsuranceRenewalCustomerView?compcode=${encodedCompCode}&utd=${encodedUTD}`;

          const customerName = wRow.CUST_NAME || "Valued Customer";
          const vehicleNo = wRow.VEHICAL_REG_NO || "N/A";
          const policyName = wRow.POLICY_NAME || "Insurance Provider";

          const waParams = [
            { type: "text", text: customerName },
            { type: "text", text: vehicleNo },
            { type: "text", text: policyName },
            { type: "text", text: customerWebLink },
          ];

          const waResult = await SendWhatsAppMessgae(
            compcode,
            wRow.CUST_MOB_NO,
            "insurance_template",
            waParams
          );

          console.log("[AUTO-WA] WhatsApp sent after call completion:", {
            compcode,
            mstUtd,
            vehicleNo,
            mobile: wRow.CUST_MOB_NO,
            questionsCount,
            waResult,
          });

        } catch (waErr) {
          console.error("[AUTO-WA] WhatsApp trigger failed:", waErr?.message || waErr);
        } finally {
          if (waSeq) {
            try { await waSeq.close(); } catch (_) { }
          }
        }
      });
      whatsappTriggered = true;
    }

    return res.status(200).json({
      Status: true,
      Message: "Webhook processed successfully",
      callId,
      saved: { status, duration, phone_number, campaign_id },
      transcript: transcriptRaw ?? null,
      summary: summary ?? null,
      insights: insights ?? null,
      whatsappTriggered,
    });
  } catch (err) {
    console.error(
      "[WEBHOOK] ERROR:",
      err?.original?.message || err?.message || err,
    );
    return res.status(500).json({
      Status: false,
      Message: err?.original?.message || err?.message || "Webhook error",
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};

exports.getCallingCustomers = async function (req, res) {
  let sequelize;
  try {
    sequelize = await dbname(req, req.headers.compcode);

    const { QueryTypes } = require("sequelize");

    const { search = "", page = 1, pageSize = 50 } = req.body || {};
    const offset = (Number(page) - 1) * Number(pageSize);
    const s = String(search || "").trim();

    // ✅ loc_code from body OR header dono support
    const loc_code =
      req.body?.loc_code ||
      req.headers?.loc_code ||
      req.headers?.["loc_code"] ||
      null;

    if (!loc_code) {
      return res
        .status(400)
        .json({ ok: false, Message: "loc_code is required" });
    }

    const locCodeArray = Array.isArray(loc_code)
      ? loc_code.map((x) => String(x).trim()).filter(Boolean)
      : String(loc_code)
        .split(",")
        .map((x) => x.trim())
        .filter(Boolean);

    if (!locCodeArray.length) {
      return res
        .status(400)
        .json({ ok: false, Message: "loc_code must not be empty" });
    }

    // ✅ COUNT query
    const countSql = `
      WITH cust AS (
        SELECT
          r.UTD,
          r.CUST_MOB_NO,
          r.CUST_NAME,
          mst.VEHICAL_REG_NO,
          CAST(r.POLICY_NUMBER AS varchar(100)) AS POLICY_NUMBER,
          ROW_NUMBER() OVER (
            PARTITION BY RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(r.CUST_MOB_NO,''),' ',''),'-',''),'+',''), 10)
            ORDER BY r.UTD DESC
          ) AS rn
        FROM dbo.INSU_RENEWAL r
        INNER JOIN dbo.INSU_RENEWAL_MST mst
          ON mst.UTD = r.TRAN_ID
        WHERE ISNULL(mst.EXPORT_TYPE, 0) = 1
          AND mst.LOC_CODE IN (:locCodeArray)
      )
      SELECT COUNT(*) AS total
      FROM cust c
      WHERE c.rn = 1
        AND (
          :search = ''
          OR c.CUST_NAME      LIKE '%' + :search + '%'
          OR c.CUST_MOB_NO    LIKE '%' + :search + '%'
          OR c.VEHICAL_REG_NO LIKE '%' + :search + '%'
          OR c.POLICY_NUMBER  LIKE '%' + :search + '%'
        )
    `;

    const countResult = await sequelize.query(countSql, {
      replacements: { search: s, locCodeArray },
      type: QueryTypes.SELECT,
    });

    const totalCount = countResult?.[0]?.total ?? 0;
    const totalPages = Math.ceil(totalCount / Number(pageSize));

    // ✅ MAIN DATA query
    const sql = `
      WITH cust AS (
        SELECT
          r.UTD,
          r.CUST_NAME,
          r.CUST_MOB_NO,
          r.POLICY_NAME,
          CAST(r.POLICY_NUMBER AS varchar(100)) AS POLICY_NUMBER,
          r.MODEL_NAME,
          CONVERT(varchar(10), r.POLICY_END_DATE, 23) AS POLICY_END_DATE,
          mst.VEHICAL_REG_NO,

          ROW_NUMBER() OVER (
            PARTITION BY RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(r.CUST_MOB_NO,''),' ',''),'-',''),'+',''), 10)
            ORDER BY r.UTD DESC
          ) AS rn
        FROM dbo.INSU_RENEWAL r
        INNER JOIN dbo.INSU_RENEWAL_MST mst
          ON mst.UTD = r.TRAN_ID
        WHERE ISNULL(mst.EXPORT_TYPE, 0) = 1
          AND mst.LOC_CODE IN (:locCodeArray)
      ),
      lastcall AS (
        SELECT
          RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(cid.mob_no,''),' ',''),'-',''),'+',''), 10) AS mob10,
          w.call_id,
          w.status,
          w.start_time,
          w.triggered_at,
          w.created_at,
          w.duration,
          w.recording_url,
          w.summary,

          ROW_NUMBER() OVER (
            PARTITION BY RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(cid.mob_no,''),' ',''),'-',''),'+',''), 10)
            ORDER BY COALESCE(w.created_at, w.triggered_at, w.start_time) DESC
          ) AS rn
        FROM dbo.call_Id_dtl cid
        LEFT JOIN dbo.call_webhook_dtl w
          ON w.call_id = cid.call_id
        WHERE cid.call_type = 'INSURANCE_RENEWAL'
      )
      SELECT
        c.UTD,
        c.CUST_NAME,
        c.CUST_MOB_NO,
        c.VEHICAL_REG_NO,
        c.POLICY_NAME,
        c.POLICY_NUMBER,
        c.MODEL_NAME,
        c.POLICY_END_DATE,

        lc.call_id        AS LAST_CALL_ID,
        lc.status         AS LAST_CALL_STATUS,
        lc.start_time     AS LAST_CALL_START_TIME,
        lc.duration       AS LAST_CALL_DURATION,
        lc.recording_url  AS LAST_CALL_RECORDING_URL,
        lc.summary        AS LAST_CALL_SUMMARY
      FROM cust c
      LEFT JOIN lastcall lc
        ON lc.mob10 = RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(c.CUST_MOB_NO,''),' ',''),'-',''),'+',''), 10)
       AND lc.rn = 1
      WHERE c.rn = 1
        AND (
          :search = ''
          OR c.CUST_NAME      LIKE '%' + :search + '%'
          OR c.CUST_MOB_NO    LIKE '%' + :search + '%'
          OR c.VEHICAL_REG_NO LIKE '%' + :search + '%'
          OR c.POLICY_NUMBER  LIKE '%' + :search + '%'
        )
      ORDER BY COALESCE(lc.created_at, lc.triggered_at, lc.start_time) DESC
      OFFSET :offset ROWS FETCH NEXT :pageSize ROWS ONLY;
    `;

    const data = await sequelize.query(sql, {
      replacements: {
        search: s,
        offset,
        pageSize: Number(pageSize),
        locCodeArray,
      },
      type: QueryTypes.SELECT,
    });

    return res.status(200).json({
      ok: true,
      data,
      page: Number(page),
      pageSize: Number(pageSize),
      totalCount,
      totalPages,
    });
  } catch (err) {
    console.error("getCallingCustomers error:", err);
    return res
      .status(500)
      .json({ ok: false, Message: err?.message || "Server error" });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};

// controllers/callingDashboard.js
exports.getCallingHistoryByMobile = async function (req, res) {
  let sequelize;
  try {
    sequelize = await dbname(req, req.headers.compcode);

    const b = req.body || {};
    const mobRaw = b.mob_no || b.phone_number || b.customerNumber || b.CUST_MOB_NO || "";
    const vehRaw = b.vehicle_number || b.VEHICAL_REG_NO || b.vehicleNo || b.RegNo || "";

    const mob10 = mobRaw ? String(mobRaw).replace(/\D/g, "").slice(-10) : null;
    const cleanVeh = vehRaw
      ? String(vehRaw).toUpperCase().replace(/[\s\-\/]+/g, "").trim()
      : null;

    if (!mob10 && !cleanVeh) {
      return res.status(400).json({ ok: false, Message: "mob_no or vehicle_number required" });
    }

    let findIdsSql = "";
    let replacements = {};

    if (cleanVeh) {
      // Jab vehicle_number available ho -> Sirf is vehicle se linked calls hi aayengi
      findIdsSql = `
        ;WITH AllCallIds AS (
          -- 1) Calls linked in FOLLOWUP_DETAILS for this vehicle via INSU_RENEWAL_MST
          SELECT DISTINCT fd.call_id
          FROM dbo.FOLLOWUP_DETAILS fd
          INNER JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = fd.TRAN_ID
          WHERE fd.call_id IS NOT NULL AND LTRIM(RTRIM(fd.call_id)) <> ''
            AND REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(m.VEHICAL_REG_NO,'')))),' ',''),'-',''),'/','') = :cleanVeh

          UNION

          -- 2) Calls linked in FOLLOWUP_DETAILS for this vehicle via INSU_RENEWAL
          SELECT DISTINCT fd.call_id
          FROM dbo.FOLLOWUP_DETAILS fd
          INNER JOIN dbo.INSU_RENEWAL r ON (r.TRAN_ID = fd.TRAN_ID OR r.UTD = fd.TRAN_ID)
          WHERE fd.call_id IS NOT NULL AND LTRIM(RTRIM(fd.call_id)) <> ''
            AND REPLACE(REPLACE(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(r.VEHICAL_REG_NO,'')))),' ',''),'-',''),'/','') = :cleanVeh
        )
        SELECT DISTINCT call_id FROM AllCallIds WHERE call_id IS NOT NULL AND LTRIM(RTRIM(call_id)) <> '';
      `;
      replacements = { cleanVeh };
    } else {
      // Jab sirf mobile number ho (vehicle_number na ho) -> Mobile number se saari calls aayengi
      findIdsSql = `
        ;WITH AllCallIds AS (
          -- 1) Calls in call_Id_dtl for this mobile (manual trigger, renewal, followup, schedule)
          SELECT DISTINCT cid.call_id
          FROM dbo.call_Id_dtl cid
          WHERE cid.call_id IS NOT NULL AND LTRIM(RTRIM(cid.call_id)) <> ''
            AND RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(cid.mob_no,''),' ',''),'-',''),'+',''), 10) = :mob10

          UNION

          -- 2) Calls directly in call_webhook_dtl for this phone number
          SELECT DISTINCT w.call_id
          FROM dbo.call_webhook_dtl w
          WHERE w.call_id IS NOT NULL AND LTRIM(RTRIM(w.call_id)) <> ''
            AND RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(w.phone_number,''),' ',''),'-',''),'+',''), 10) = :mob10

          UNION

          -- 3) Calls in FOLLOWUP_DETAILS for customer's mobile
          SELECT DISTINCT fd.call_id
          FROM dbo.FOLLOWUP_DETAILS fd
          INNER JOIN dbo.INSU_RENEWAL r ON (r.TRAN_ID = fd.TRAN_ID OR r.UTD = fd.TRAN_ID)
          WHERE fd.call_id IS NOT NULL AND LTRIM(RTRIM(fd.call_id)) <> ''
            AND RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(r.CUST_MOB_NO,''),' ',''),'-',''),'+',''), 10) = :mob10
        )
        SELECT DISTINCT call_id FROM AllCallIds WHERE call_id IS NOT NULL AND LTRIM(RTRIM(call_id)) <> '';
      `;
      replacements = { mob10 };
    }

    const idRows = await sequelize.query(findIdsSql, {
      replacements,
      type: QueryTypes.SELECT,
    });

    const allCallIds = (idRows || []).map((r) => r.call_id).filter(Boolean);

    if (!allCallIds.length) {
      return res.status(200).json({ ok: true, data: [] });
    }

    // Auto-sync any call that is still INITIATED or missing transcript/summary/recording
    for (const cId of allCallIds) {
      try {
        const existing = await sequelize.query(
          `SELECT TOP 1 status, transcript, summary, duration, recording_url FROM dbo.call_webhook_dtl WHERE call_id = :callId`,
          { replacements: { callId: cId }, type: QueryTypes.SELECT }
        );
        const st = String(existing?.[0]?.status || '').toUpperCase();
        const isTerminal = ['COMPLETED', 'ENDED', 'FINISHED', 'FAILED', 'BUSY', 'NO_ANSWER', 'CANCELED'].includes(st);
        const isMissingSummary = ['COMPLETED', 'ENDED', 'FINISHED'].includes(st) && (!existing?.[0]?.summary || String(existing[0].summary).trim() === '' || String(existing[0].summary).includes('"actionType"') || String(existing[0].summary).startsWith('{'));

        if (!existing?.length || !isTerminal || isMissingSummary) {
          const apiRes = await getCallStatus1(cId);
          if (apiRes) {
            await saveCallWebhookDetails(sequelize, cId, apiRes);
          }
        }
      } catch (e) {
        console.error("[CALLING HISTORY SYNC] Error fetching status for callId:", cId, e?.message);
      }
    }

    const sql = `
      ;WITH RankedCalls AS (
        SELECT
          w.call_id,
          w.campaign_id,
          w.direction,
          w.phone_number,
          w.status,
          CONVERT(varchar(19), w.triggered_at, 120) AS triggered_at,
          CONVERT(varchar(19), w.start_time, 120)   AS start_time,
          CONVERT(varchar(19), w.end_time, 120)     AS end_time,
          w.duration,
          w.ring_time,
          w.summary,
          w.transcript,
          w.callee_name,
          w.category,
          w.recording_url,
          CONVERT(varchar(19), w.created_at, 120)   AS created_at,
          w.created_at                              AS raw_created_at,
          w.start_time                              AS raw_start_time,
          w.triggered_at                            AS raw_triggered_at,
          ROW_NUMBER() OVER (
            PARTITION BY LTRIM(RTRIM(w.call_id))
            ORDER BY 
              CASE WHEN UPPER(ISNULL(w.status,'')) IN ('COMPLETED','ENDED','FINISHED') THEN 1 ELSE 2 END,
              COALESCE(w.end_time, w.start_time, w.triggered_at, w.created_at) DESC,
              w.created_at DESC
          ) AS rn
        FROM dbo.call_webhook_dtl w
        WHERE w.call_id IN (:allCallIds)
      )
      SELECT TOP 200
        call_id,
        campaign_id,
        direction,
        phone_number,
        status,
        triggered_at,
        start_time,
        end_time,
        duration,
        ring_time,
        summary,
        transcript,
        callee_name,
        category,
        recording_url,
        created_at
      FROM RankedCalls
      WHERE rn = 1
      ORDER BY COALESCE(raw_start_time, raw_triggered_at, raw_created_at) DESC, raw_created_at DESC;
    `;

    const data = await sequelize.query(sql, {
      replacements: { allCallIds },
      type: QueryTypes.SELECT,
    });

    const cleanedData = (data || []).map((row) => ({
      ...row,
      summary: extractCleanSummary(row.summary),
    }));

    return res.status(200).json({ ok: true, data: cleanedData });
  } catch (err) {
    return res
      .status(500)
      .json({ ok: false, Message: err?.message || "Server error" });
  } finally {
    if (sequelize)
      try {
        await sequelize.close();
      } catch (_) { }
  }
};

exports.streamCallRecordingByCallId = async function (req, res) {
  let sequelize;
  try {
    const { callId } = req.params;
    if (!callId) return res.status(400).send("callId required");

    sequelize = await dbname(req, req.headers.compcode);

    // ✅ optional security: check callId belongs to mob_no (if provided)
    const mobNo = req.query.mob_no ? String(req.query.mob_no) : "";
    const mob10 = mobNo ? mobNo.replace(/\D/g, "").slice(-10) : "";

    if (mob10) {
      const checkSql = `
        SELECT TOP 1 cid.call_id
        FROM dbo.call_Id_dtl cid
        WHERE cid.call_type = 'INSURANCE_RENEWAL'
          AND cid.call_id = :callId
          AND RIGHT(REPLACE(REPLACE(REPLACE(ISNULL(cid.mob_no,''),' ',''),'-',''),'+',''), 10) = :mob10
      `;
      const okRow = await sequelize.query(checkSql, {
        replacements: { callId, mob10 },
        type: QueryTypes.SELECT,
      });

      if (!okRow?.length) {
        return res.status(403).json({
          ok: false,
          Message: "This callId does not belong to given mob_no",
        });
      }
    }

    // ✅ stream from Callmatic (API key server side)
    await getCallRecording1(callId, res);
  } catch (err) {
    console.error("streamCallRecordingByCallId error:", err?.message);
    return res.status(500).send(err?.message || "Recording stream failed");
  } finally {
    if (sequelize)
      try {
        await sequelize.close();
      } catch (_) { }
  }
};

exports.createInsuCallingConfig = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);

  try {
    const b = req.body || {};

    const INSU_COMPANY_NAME =
      b.INSU_COMPANY_NAME != null ? String(b.INSU_COMPANY_NAME).trim() : "";

    const SLOT1 = b.SLOT1 != null ? String(b.SLOT1).trim() : "";
    const SLOT2 = b.SLOT2 != null ? String(b.SLOT2).trim() : "";
    const SLOT3 = b.SLOT3 != null ? String(b.SLOT3).trim() : "";

    if (!INSU_COMPANY_NAME) {
      return res
        .status(400)
        .json({ success: false, message: "INSU_COMPANY_NAME is required" });
    }
    if (!SLOT1 || !SLOT2 || !SLOT3) {
      return res
        .status(400)
        .json({ success: false, message: "SLOT1, SLOT2, SLOT3 are required" });
    }

    const toTimeHHMMSS = (v) => {
      if (v === null || v === undefined || String(v).trim() === "") return null;
      const s = String(v).trim();

      // 24h: HH:MM or HH:MM:SS
      let m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (m) {
        const hh = String(
          Math.min(23, Math.max(0, parseInt(m[1], 10))),
        ).padStart(2, "0");
        const mm = String(
          Math.min(59, Math.max(0, parseInt(m[2], 10))),
        ).padStart(2, "0");
        const ss = String(
          m[3] != null ? Math.min(59, Math.max(0, parseInt(m[3], 10))) : 0,
        ).padStart(2, "0");
        return `${hh}:${mm}:${ss}`;
      }

      // 12h: h:mm AM/PM
      m = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
      if (m) {
        let h = parseInt(m[1], 10);
        const min = parseInt(m[2], 10);
        const ap = m[3].toUpperCase();

        if (h < 1 || h > 12 || min < 0 || min > 59) return null;

        if (ap === "PM" && h !== 12) h += 12;
        if (ap === "AM" && h === 12) h = 0;

        return `${String(h).padStart(2, "0")}:${String(min).padStart(2, "0")}:00`;
      }

      return null;
    };

    const SALES_EXECUTIVE_NO =
      b.SALES_EXECUTIVE_NO !== undefined &&
        b.SALES_EXECUTIVE_NO !== null &&
        String(b.SALES_EXECUTIVE_NO).trim() !== ""
        ? Number(b.SALES_EXECUTIVE_NO)
        : null;

    const CAMPAIGN_ID =
      b.CAMPAIGN_ID != null && String(b.CAMPAIGN_ID).trim() !== ""
        ? String(b.CAMPAIGN_ID).trim()
        : null;

    const CALLBACK_TIME = toTimeHHMMSS(b.CALLBACK_TIME);

    if (
      b.CALLBACK_TIME != null &&
      String(b.CALLBACK_TIME).trim() !== "" &&
      !CALLBACK_TIME
    ) {
      return res.status(400).json({
        success: false,
        message: "Invalid CALLBACK_TIME. Use 'HH:MM:SS' (24h) or 'h:mm AM/PM'.",
        example: ["19:24:00", "7:24 PM"],
      });
    }

    const LOC_CODE =
      b.LOC_CODE === null ||
        b.LOC_CODE === undefined ||
        String(b.LOC_CODE).trim() === ""
        ? null
        : Number(b.LOC_CODE);

    const STATUS =
      b.STATUS != null && String(b.STATUS).trim() !== ""
        ? String(b.STATUS).trim()
        : null;

    const CREATED_BY =
      b.CREATED_BY != null && String(b.CREATED_BY).trim() !== ""
        ? String(b.CREATED_BY).trim()
        : null;

    const insertSQL = `
      INSERT INTO dbo.INSU_CALLING_CONFIG
      (
        INSU_COMPANY_NAME,
        SALES_EXECUTIVE_NO,
        SLOT1,
        SLOT2,
        SLOT3,
        CAMPAIGN_ID,
        CALLBACK_TIME,
        LOC_CODE,
        STATUS,
        CREATED_BY
      )
      OUTPUT
        INSERTED.UTD,
        INSERTED.INSU_COMPANY_NAME,
        INSERTED.SALES_EXECUTIVE_NO,
        INSERTED.SLOT1,
        INSERTED.SLOT2,
        INSERTED.SLOT3,
        INSERTED.CAMPAIGN_ID,
        CONVERT(varchar(8), INSERTED.CALLBACK_TIME) AS CALLBACK_TIME,
        INSERTED.LOC_CODE,
        INSERTED.STATUS,
        INSERTED.CREATED_BY,
        CONVERT(varchar(19), INSERTED.CREATED_AT, 120) AS CREATED_AT
      VALUES
      (
        :INSU_COMPANY_NAME,
        :SALES_EXECUTIVE_NO,
        :SLOT1,
        :SLOT2,
        :SLOT3,
        :CAMPAIGN_ID,
        :CALLBACK_TIME,
        :LOC_CODE,
        :STATUS,
        :CREATED_BY
      )
    `;

    const [inserted] = await sequelize.query(insertSQL, {
      type: sequelize.QueryTypes.SELECT,
      replacements: {
        INSU_COMPANY_NAME,
        SALES_EXECUTIVE_NO: Number.isFinite(SALES_EXECUTIVE_NO)
          ? SALES_EXECUTIVE_NO
          : null,
        SLOT1,
        SLOT2,
        SLOT3,
        CAMPAIGN_ID,
        CALLBACK_TIME,
        LOC_CODE: Number.isFinite(LOC_CODE) ? LOC_CODE : null,
        STATUS,
        CREATED_BY,
      },
    });

    return res.status(201).json({
      success: true,
      message: "INSU_CALLING_CONFIG inserted successfully",
      data: inserted || null,
    });
  } catch (err) {
    console.error("[INSU_CALLING_CONFIG] INSERT | Error:", err);
    return res
      .status(500)
      .json({ success: false, message: err?.message || "Server error" });
  } finally {
    await sequelize.close();
  }
};

exports.getInsuCallingConfig = async (req, res) => {
  let sequelize;
  try {
    sequelize = await dbname(req, req.headers.compcode);

    const { status, Loc_Code, loc_code, INSU_COMPANY_NAME } = req.query;

    const whereClauses = [];
    const replacements = {};

    // Optional: filter by company name (exact match)
    if (
      INSU_COMPANY_NAME !== undefined &&
      String(INSU_COMPANY_NAME).trim() !== ""
    ) {
      whereClauses.push("c.INSU_COMPANY_NAME = :insuCompanyName");
      replacements.insuCompanyName = String(INSU_COMPANY_NAME).trim();
    }

    // ✅ Status filter (reference style)
    // If not provided or "all" => return all.
    // Table column is varchar, so we support:
    // - status=1 -> matches '1' OR 'ACTIVE'
    // - status=0 -> matches '0' OR 'INACTIVE'
    // - status='ACTIVE'/'INACTIVE' -> exact
    if (
      status !== undefined &&
      status !== "" &&
      String(status).toLowerCase() !== "all"
    ) {
      const s = String(status).trim();

      // numeric status support
      if (s === "1" || s === "0") {
        const statusText = s === "1" ? "ACTIVE" : "INACTIVE";
        whereClauses.push(
          "(ISNULL(c.STATUS,'') = :statusRaw OR ISNULL(c.STATUS,'') = :statusText)",
        );
        replacements.statusRaw = s; // '1' or '0'
        replacements.statusText = statusText; // 'ACTIVE' or 'INACTIVE'
      } else {
        whereClauses.push("ISNULL(c.STATUS,'') = :statusExact");
        replacements.statusExact = s;
      }
    }

    // ✅ LOC_CODE filter (single)
    const locVal = Loc_Code ?? loc_code;
    if (
      locVal !== undefined &&
      locVal !== null &&
      String(locVal).trim() !== ""
    ) {
      whereClauses.push("c.LOC_CODE = :Loc_Code");
      replacements.Loc_Code = Number(locVal);
    }

    const whereSQL = whereClauses.length
      ? `WHERE ${whereClauses.join(" AND ")}`
      : "";

    const rows = await sequelize.query(
      `
      SELECT
        c.UTD,
        c.LOC_CODE,
        mm.Misc_Name                            AS LOC_NAME,     -- optional (if Misc_Type=85 exists for branches)
        c.INSU_COMPANY_NAME,
        c.SALES_EXECUTIVE_NO,
        c.SLOT1,
        c.SLOT2,
        c.SLOT3,
        CONVERT(varchar(8), c.CALLBACK_TIME)    AS CALLBACK_TIME,
        c.CAMPAIGN_ID,
        c.STATUS,
        c.CREATED_BY,
        CONVERT(varchar(19), c.CREATED_AT, 120) AS CREATED_AT
      FROM dbo.INSU_CALLING_CONFIG c
      LEFT JOIN dbo.MISC_MST mm
        ON mm.Misc_Code = CAST(c.LOC_CODE AS NVARCHAR(20))
       AND mm.Misc_Type = 85
      ${whereSQL}
      ORDER BY c.UTD DESC
      `,
      { replacements, type: QueryTypes.SELECT },
    );

    return res.status(200).json({
      success: true,
      totalRecords: rows.length,
      data: rows,
    });
  } catch (err) {
    console.error("[INSU_CALLING_CONFIG] GET ALL | Error:", err?.message);
    return res.status(500).json({ success: false, message: err?.message });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};

exports.updateInsuCallingConfig = async (req, res) => {
  let sequelize;
  try {
    sequelize = await dbname(req, req.headers.compcode);

    const b = req.body || {};

    // Accept both key styles
    const UTD = b.UTD;
    const Loc_Code = b.Loc_Code ?? b.LOC_CODE ?? b.loc_code;

    const INSU_COMPANY_NAME = b.INSU_COMPANY_NAME;
    const SALES_EXECUTIVE_NO = b.SALES_EXECUTIVE_NO;

    const SLOT1 = b.SLOT1;
    const SLOT2 = b.SLOT2;
    const SLOT3 = b.SLOT3;

    const CALLBACK_TIME = b.CALLBACK_TIME;
    const CAMPAIGN_ID = b.CAMPAIGN_ID;

    const STATUS = b.STATUS; // varchar in DB
    const CREATED_BY = b.Created_By ?? b.CREATED_BY ?? b.created_by; // table has CREATED_BY only

    if (!UTD) {
      return res
        .status(400)
        .json({ success: false, message: "UTD is required for update" });
    }

    if (!INSU_COMPANY_NAME || String(INSU_COMPANY_NAME).trim() === "") {
      return res.status(400).json({
        success: false,
        message: "INSU_COMPANY_NAME is required",
      });
    }

    // SLOT1/2/3 are NOT NULL in your table -> must be provided (otherwise update can fail)
    if (!SLOT1 || !SLOT2 || !SLOT3) {
      return res.status(400).json({
        success: false,
        message: "SLOT1, SLOT2, SLOT3 are required",
      });
    }

    // Normalize CALLBACK_TIME to HH:MM:SS (or null)
    const toTimeHHMMSS = (v) => {
      if (v === null || v === undefined || String(v).trim() === "") return null;
      const s = String(v).trim();

      // 24h: HH:MM or HH:MM:SS
      let m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (m) {
        const hh = String(
          Math.min(23, Math.max(0, parseInt(m[1], 10))),
        ).padStart(2, "0");
        const mm = String(
          Math.min(59, Math.max(0, parseInt(m[2], 10))),
        ).padStart(2, "0");
        const ss = String(
          m[3] != null ? Math.min(59, Math.max(0, parseInt(m[3], 10))) : 0,
        ).padStart(2, "0");
        return `${hh}:${mm}:${ss}`;
      }

      // 12h: h:mm AM/PM
      m = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
      if (m) {
        let h = parseInt(m[1], 10);
        const min = parseInt(m[2], 10);
        const ap = m[3].toUpperCase();

        if (h < 1 || h > 12 || min < 0 || min > 59) return null;

        if (ap === "PM" && h !== 12) h += 12;
        if (ap === "AM" && h === 12) h = 0;

        return `${String(h).padStart(2, "0")}:${String(min).padStart(2, "0")}:00`;
      }

      return null;
    };

    const callbackTimeFinal = toTimeHHMMSS(CALLBACK_TIME);
    if (
      CALLBACK_TIME != null &&
      String(CALLBACK_TIME).trim() !== "" &&
      !callbackTimeFinal
    ) {
      return res.status(400).json({
        success: false,
        message: "Invalid CALLBACK_TIME. Use 'HH:MM:SS' (24h) or 'h:mm AM/PM'.",
        example: ["19:24:00", "7:24 PM"],
      });
    }

    const locCodeFinal =
      Loc_Code === null ||
        Loc_Code === undefined ||
        String(Loc_Code).trim() === ""
        ? null
        : Number(Loc_Code);

    const salesExecFinal =
      SALES_EXECUTIVE_NO === null ||
        SALES_EXECUTIVE_NO === undefined ||
        String(SALES_EXECUTIVE_NO).trim() === ""
        ? null
        : Number(SALES_EXECUTIVE_NO);

    const statusFinal =
      STATUS === null || STATUS === undefined || String(STATUS).trim() === ""
        ? null
        : String(STATUS).trim();

    const campaignFinal =
      CAMPAIGN_ID === null ||
        CAMPAIGN_ID === undefined ||
        String(CAMPAIGN_ID).trim() === ""
        ? null
        : String(CAMPAIGN_ID).trim();

    // ✅ FIX: UPDATED_BY / UPDATED_AT removed (columns not in your table)
    const updateSQL = `
      UPDATE dbo.INSU_CALLING_CONFIG
      SET
        LOC_CODE           = :LOC_CODE,
        INSU_COMPANY_NAME  = :INSU_COMPANY_NAME,
        SALES_EXECUTIVE_NO = :SALES_EXECUTIVE_NO,
        SLOT1              = :SLOT1,
        SLOT2              = :SLOT2,
        SLOT3              = :SLOT3,
        CALLBACK_TIME      = :CALLBACK_TIME,
        CAMPAIGN_ID        = :CAMPAIGN_ID,
        STATUS             = :STATUS,
        CREATED_BY         = COALESCE(:CREATED_BY, CREATED_BY)
      OUTPUT
        INSERTED.UTD,
        INSERTED.LOC_CODE,
        INSERTED.INSU_COMPANY_NAME,
        INSERTED.SALES_EXECUTIVE_NO,
        INSERTED.SLOT1,
        INSERTED.SLOT2,
        INSERTED.SLOT3,
        CONVERT(varchar(8), INSERTED.CALLBACK_TIME) AS CALLBACK_TIME,
        INSERTED.CAMPAIGN_ID,
        INSERTED.STATUS,
        INSERTED.CREATED_BY,
        CONVERT(varchar(19), INSERTED.CREATED_AT, 120) AS CREATED_AT
      WHERE UTD = :UTD
    `;

    const updatedRows = await sequelize.query(updateSQL, {
      replacements: {
        UTD: Number(UTD),
        LOC_CODE: Number.isFinite(locCodeFinal) ? locCodeFinal : null,
        INSU_COMPANY_NAME: String(INSU_COMPANY_NAME).trim(),
        SALES_EXECUTIVE_NO: Number.isFinite(salesExecFinal)
          ? salesExecFinal
          : null,
        SLOT1: String(SLOT1).trim(),
        SLOT2: String(SLOT2).trim(),
        SLOT3: String(SLOT3).trim(),
        CALLBACK_TIME: callbackTimeFinal, // HH:MM:SS or null
        CAMPAIGN_ID: campaignFinal,
        STATUS: statusFinal,
        CREATED_BY:
          CREATED_BY != null && String(CREATED_BY).trim() !== ""
            ? String(CREATED_BY).trim()
            : null,
      },
      type: QueryTypes.SELECT,
    });

    if (!updatedRows || !updatedRows.length) {
      return res.status(404).json({
        success: false,
        message: "Record not found for given UTD",
      });
    }

    return res.status(200).json({
      success: true,
      message: "Insurance Calling Config updated successfully",
      data: updatedRows[0],
    });
  } catch (err) {
    console.error("[INSU-CONFIG] updateInsuCallingConfig error:", err);
    return res.status(500).json({
      success: false,
      message: err?.message || "Server error",
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};

exports.toggleInsuCallingConfigStatus = async (req, res) => {
  let sequelize;
  try {
    sequelize = await dbname(req, req.headers.compcode);

    const { UTD, status } = req.body;

    if (!UTD) {
      return res
        .status(400)
        .json({ success: false, message: "UTD is required" });
    }

    // STATUS column is varchar(50) in your table, and UPDATED_AT column does NOT exist.
    // So: update only STATUS.
    // Accept status as: 1/0 or ACTIVE/INACTIVE
    let statusFinal = null;

    if (
      status === null ||
      status === undefined ||
      String(status).trim() === ""
    ) {
      return res
        .status(400)
        .json({ success: false, message: "status is required" });
    }

    const s = String(status).trim();

    if (s === "1") statusFinal = "ACTIVE";
    else if (s === "0") statusFinal = "INACTIVE";
    else statusFinal = s; // allow 'ACTIVE'/'INACTIVE'/'SOMETHING'

    const sql = `
      UPDATE dbo.INSU_CALLING_CONFIG
      SET STATUS = :status
      OUTPUT
        INSERTED.UTD,
        INSERTED.STATUS,
        INSERTED.LOC_CODE,
        INSERTED.INSU_COMPANY_NAME
      WHERE UTD = :UTD
    `;

    const rows = await sequelize.query(sql, {
      replacements: { UTD: Number(UTD), status: statusFinal },
      type: QueryTypes.SELECT,
    });

    if (!rows || !rows.length) {
      return res.status(404).json({
        success: false,
        message: "Record not found for given UTD",
      });
    }

    return res.status(200).json({
      success: true,
      message: `Config status updated successfully`,
      data: rows[0],
    });
  } catch (err) {
    console.error("[INSU-CONFIG] toggleInsuCallingConfigStatus error:", err);
    return res.status(500).json({
      success: false,
      message: err?.message || "Server error",
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};

exports.getInsuranceDashboardMetrics = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  const { QueryTypes } = require("sequelize");

  try {
    const b = req.body || {};

    // ─────────────────────────────────────────────────────────────
    // 1. HELPERS
    // ─────────────────────────────────────────────────────────────
    const pad2 = (x) => String(x).padStart(2, "0");

    const toDbDate = (v) => {
      if (!v) return null;
      const s = String(v).trim();
      let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
      m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      return null;
    };

    // ─────────────────────────────────────────────────────────────
    // 2. FILTERS
    // ─────────────────────────────────────────────────────────────
    const fromDate = toDbDate(b.fromDate);
    const toDate = toDbDate(b.toDate);

    console.log("[Dashboard] Raw dates from request:", {
      fromDateRaw: b.fromDate,
      toDateRaw: b.toDate,
      fromDateParsed: fromDate,
      toDateParsed: toDate,
    });

    const locRaw =
      b.loc_code !== undefined &&
        b.loc_code !== null &&
        String(b.loc_code).trim() !== ""
        ? b.loc_code
        : req.headers.loc_code !== undefined &&
          req.headers.loc_code !== null &&
          String(req.headers.loc_code).trim() !== ""
          ? req.headers.loc_code
          : null;

    let locCodes = [];
    if (Array.isArray(locRaw)) {
      locCodes = locRaw
        .map((x) => Number(String(x).trim()))
        .filter((n) => Number.isFinite(n));
    } else if (locRaw != null && String(locRaw).trim() !== "") {
      const str = String(locRaw).trim();
      if (str.includes(",")) {
        locCodes = str
          .split(",")
          .map((x) => Number(String(x).trim()))
          .filter((n) => Number.isFinite(n));
      } else {
        const n = Number(str);
        if (Number.isFinite(n)) locCodes = [n];
      }
    }

    console.log("[Dashboard] loc_code raw:", locRaw, "→ parsed:", locCodes);

    const exportType = parseInt(b.EXPORT_TYPE ?? 1, 10);
    const exportTypeFinal = Number.isFinite(exportType) ? exportType : 1;

    const periodType = String(b.periodType || "monthly")
      .toLowerCase()
      .trim();
    const validPeriods = ["daily", "weekly", "monthly", "yearly"];
    const periodTypeFinal = validPeriods.includes(periodType)
      ? periodType
      : "monthly";

    // ─────────────────────────────────────────────────────────────
    // PERIOD CONFIG
    // ─────────────────────────────────────────────────────────────

    let periodFormatExpr = "";
    let periodRangeStart = "";
    let periodRangeEnd = "";
    let periodOrderExpr = "";
    let periodLabel = "";

    if (periodTypeFinal === "daily") {
      periodFormatExpr = `FORMAT(m.CREATED_AT, 'yyyy-MM-dd')`;
      periodRangeStart = `CAST(GETDATE() AS date)`;
      periodRangeEnd = `CAST(GETDATE() AS date)`;
      periodOrderExpr = `FORMAT(m.CREATED_AT, 'yyyy-MM-dd')`;
      periodLabel = "day";
    } else if (periodTypeFinal === "weekly") {
      periodFormatExpr = `CONCAT(YEAR(m.CREATED_AT), '-W', RIGHT('0' + CAST(DATEPART(iso_week, m.CREATED_AT) AS VARCHAR(2)), 2))`;
      periodRangeStart = `DATEADD(day, 2 - DATEPART(weekday, GETDATE()), CAST(GETDATE() AS date))`;
      periodRangeEnd = `DATEADD(day, 8 - DATEPART(weekday, GETDATE()), CAST(GETDATE() AS date))`;
      periodOrderExpr = `MIN(m.CREATED_AT)`;
      periodLabel = "week";
    } else if (periodTypeFinal === "yearly") {
      periodFormatExpr = `FORMAT(m.CREATED_AT, 'yyyy')`;
      periodRangeStart = `DATEFROMPARTS(YEAR(GETDATE()), 1, 1)`;
      periodRangeEnd = `DATEFROMPARTS(YEAR(GETDATE()), 12, 31)`;
      periodOrderExpr = `FORMAT(m.CREATED_AT, 'yyyy')`;
      periodLabel = "year";
    } else {
      periodFormatExpr = `FORMAT(m.CREATED_AT, 'yyyy-MM')`;
      periodRangeStart = `DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)`;
      periodRangeEnd = `EOMONTH(GETDATE())`;
      periodOrderExpr = `FORMAT(m.CREATED_AT, 'yyyy-MM')`;
      periodLabel = "month";
    }

    const locFilterMSQL = locCodes.length
      ? `AND (m.loc_code IN (:locCodes) OR m.loc_code IS NULL OR m.loc_code = 0)`
      : "";
    const locFilterPSQL = locCodes.length
      ? `AND (m.loc_code IN (:locCodes) OR m.loc_code IS NULL OR m.loc_code = 0)`
      : "";

    // ✅ FIXED: सभी queries को अलग अलग filter variables दिए
    const replacements = { exportType: exportTypeFinal };
    if (locCodes.length) replacements.locCodes = locCodes;

    // For KPI, Tables, Distribution queries (based on m.CREATED_AT)
    let dateFilterSQL = "";
    let pymtDateFilterSQL = "";
    if (fromDate && toDate) {
      dateFilterSQL = ` AND CAST(m.CREATED_AT AS DATE) BETWEEN :fromDate AND :toDate`;
      pymtDateFilterSQL = ` AND (
        CAST(COALESCE(p.PYMT_DATE, p.CREATED_AT, m.CREATED_AT) AS DATE) BETWEEN :fromDate AND :toDate
        OR CAST(p.CREATED_AT AS DATE) BETWEEN :fromDate AND :toDate
        OR p.PYMT_DATE IS NULL
      )`;
      replacements.fromDate = fromDate;
      replacements.toDate = toDate;
    } else if (fromDate) {
      dateFilterSQL = ` AND CAST(m.CREATED_AT AS DATE) >= :fromDate`;
      pymtDateFilterSQL = ` AND (
        CAST(COALESCE(p.PYMT_DATE, p.CREATED_AT, m.CREATED_AT) AS DATE) >= :fromDate
        OR CAST(p.CREATED_AT AS DATE) >= :fromDate
        OR p.PYMT_DATE IS NULL
      )`;
      replacements.fromDate = fromDate;
    } else if (toDate) {
      dateFilterSQL = ` AND CAST(m.CREATED_AT AS DATE) <= :toDate`;
      pymtDateFilterSQL = ` AND (
        CAST(COALESCE(p.PYMT_DATE, p.CREATED_AT, m.CREATED_AT) AS DATE) <= :toDate
        OR CAST(p.CREATED_AT AS DATE) <= :toDate
        OR p.PYMT_DATE IS NULL
      )`;
      replacements.toDate = toDate;
    } else {
      // Default live dashboard: only current month created policies & payments
      dateFilterSQL = ` AND (
        CAST(m.CREATED_AT AS DATE) BETWEEN DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1) AND EOMONTH(GETDATE())
        OR m.CREATED_AT IS NULL
      )`;
      pymtDateFilterSQL = ` AND (
        CAST(COALESCE(p.PYMT_DATE, p.CREATED_AT, m.CREATED_AT) AS DATE) BETWEEN DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1) AND EOMONTH(GETDATE())
        OR CAST(p.CREATED_AT AS DATE) BETWEEN DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1) AND EOMONTH(GETDATE())
        OR p.PYMT_DATE IS NULL
      )`;
    }

    // Period date filter (based on m.CREATED_AT)
    const periodReplacements = { ...replacements };
    let periodDateFilterSQL = "";

    if (fromDate && toDate) {
      periodDateFilterSQL = ` AND CAST(m.CREATED_AT AS DATE) BETWEEN :fromDate AND :toDate`;
      periodReplacements.fromDate = fromDate;
      periodReplacements.toDate = toDate;
    } else if (fromDate) {
      periodDateFilterSQL = ` AND CAST(m.CREATED_AT AS DATE) >= :fromDate`;
      periodReplacements.fromDate = fromDate;
    } else if (toDate) {
      periodDateFilterSQL = ` AND CAST(m.CREATED_AT AS DATE) <= :toDate`;
      periodReplacements.toDate = toDate;
    } else {
      periodDateFilterSQL = ` AND CAST(m.CREATED_AT AS DATE) BETWEEN ${periodRangeStart} AND ${periodRangeEnd}`;
    }

    // ─────────────────────────────────────────────────────────────
    // 3. READ CONFIG
    // ─────────────────────────────────────────────────────────────
    const [keyRow] = await sequelize.query(
      `SELECT TOP 1
         ISNULL(CAST(INSU_BEFORE_DAYS AS INT), 30) AS beforeDays,
         ISNULL(CAST(INSU_AFTER_DAYS  AS INT), 30) AS afterDays
       FROM dbo.COMP_KEYDATA`,
      { type: QueryTypes.SELECT }
    );

    const beforeDays = Number.isFinite(parseInt(keyRow?.beforeDays, 10))
      ? parseInt(keyRow.beforeDays, 10)
      : 30;
    const afterDays = Number.isFinite(parseInt(keyRow?.afterDays, 10))
      ? parseInt(keyRow.afterDays, 10)
      : 30;

    // ─────────────────────────────────────────────────────────────
    // 4. KPI STATS
    // ─────────────────────────────────────────────────────────────
    const policyKpiQuery = `
      SELECT
        COUNT(DISTINCT m.UTD) AS totalPolicies,

        COUNT(DISTINCT CASE
          WHEN CAST(r.POLICY_END_DATE AS DATE) >= CAST(GETDATE() AS DATE)
          THEN m.UTD END) AS activePolicies,

        COUNT(DISTINCT CASE
          WHEN r.POLICY_END_DATE IS NULL OR CAST(r.POLICY_END_DATE AS DATE) < CAST(GETDATE() AS DATE)
          THEN m.UTD END) AS expiredPolicies,

        COUNT(DISTINCT CASE
          WHEN CAST(r.POLICY_END_DATE AS DATE)
               BETWEEN CAST(GETDATE() AS DATE)
               AND DATEADD(day, 30, CAST(GETDATE() AS DATE))
          THEN m.UTD END) AS expiringIn30Days,

        ISNULL(SUM(COALESCE(NULLIF(r.PREMIUM_AMOUNT, 0), 0)), 0) AS totalPremiumAmount

      FROM dbo.INSU_RENEWAL_MST m
      OUTER APPLY (
        SELECT TOP 1
          ir.POLICY_END_DATE,
          ir.PREMIUM_AMOUNT
        FROM dbo.INSU_RENEWAL ir
        WHERE (ir.TRAN_ID = m.UTD OR (m.VEHICAL_REG_NO IS NOT NULL AND REPLACE(UPPER(ir.VEHICAL_REG_NO), ' ', '') = REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '')))
          AND ISNULL(ir.EXPORT_TYPE, 1) = :exportType
        ORDER BY ir.UTD DESC
      ) r
      WHERE ISNULL(m.EXPORT_TYPE, 1) = :exportType
        ${locFilterMSQL}
        ${dateFilterSQL}
    `;

    const [policyKpiStats] = await sequelize.query(policyKpiQuery, {
      replacements,
      type: QueryTypes.SELECT,
    });

    const paymentKpiQuery = `
      ;WITH DistinctPymt AS (
        SELECT 
          p.UTD,
          p.TRAN_ID,
          p.ACNT_APPR_STATUS,
          p.ACNT_APPR_REMARK,
          p.ACNT_APPR_CODE,
          p.ACNT_APPR_DATE,
          p.PYMT_AMOUNT,
          p.PREMIUM_AMOUNT,
          ROW_NUMBER() OVER (
            PARTITION BY COALESCE(NULLIF(LTRIM(RTRIM(m.VEHICAL_REG_NO)), ''), CAST(p.TRAN_ID AS VARCHAR))
            ORDER BY p.UTD DESC
          ) AS rn
        FROM dbo.INSU_RENEWAL_PYMT p
        LEFT JOIN dbo.INSU_RENEWAL_MST m
          ON m.UTD = p.TRAN_ID
        WHERE p.PYMT_STATUS = 1
          ${locFilterPSQL}
          ${pymtDateFilterSQL}
      )
      SELECT
        COUNT(CASE
          WHEN p.ACNT_APPR_STATUS = 1
          THEN p.UTD END) AS approvedPaymentsCount,

        COUNT(CASE
          WHEN p.ACNT_APPR_STATUS = 2 OR (p.ACNT_APPR_STATUS = 0 AND (NULLIF(p.ACNT_APPR_REMARK, '') IS NOT NULL OR NULLIF(p.ACNT_APPR_CODE, '') IS NOT NULL OR p.ACNT_APPR_DATE IS NOT NULL))
          THEN p.UTD END) AS rejectedPaymentsCount,

        COUNT(CASE
          WHEN p.ACNT_APPR_STATUS IS NULL OR (p.ACNT_APPR_STATUS = 0 AND NULLIF(p.ACNT_APPR_REMARK, '') IS NULL AND NULLIF(p.ACNT_APPR_CODE, '') IS NULL AND p.ACNT_APPR_DATE IS NULL)
          THEN p.UTD END) AS pendingPaymentsCount,

        COUNT(p.UTD) AS totalPaymentsCount,

        ISNULL(SUM(CASE
          WHEN p.ACNT_APPR_STATUS = 1
          THEN p.PYMT_AMOUNT ELSE 0 END), 0) AS totalCollectedAmount,

        ISNULL(SUM(CASE
          WHEN p.ACNT_APPR_STATUS = 2 OR (p.ACNT_APPR_STATUS = 0 AND (NULLIF(p.ACNT_APPR_REMARK, '') IS NOT NULL OR NULLIF(p.ACNT_APPR_CODE, '') IS NOT NULL OR p.ACNT_APPR_DATE IS NOT NULL))
          THEN p.PYMT_AMOUNT ELSE 0 END), 0) AS totalRejectedAmount,

        ISNULL(SUM(CASE
          WHEN p.ACNT_APPR_STATUS IS NULL OR (p.ACNT_APPR_STATUS = 0 AND NULLIF(p.ACNT_APPR_REMARK, '') IS NULL AND NULLIF(p.ACNT_APPR_CODE, '') IS NULL AND p.ACNT_APPR_DATE IS NULL)
          THEN p.PYMT_AMOUNT ELSE 0 END), 0) AS totalPendingAmount,

        ISNULL(SUM(COALESCE(NULLIF(p.PREMIUM_AMOUNT, 0), p.PYMT_AMOUNT, 0)), 0) AS totalPaymentPremium

      FROM DistinctPymt p
      WHERE p.rn = 1
    `;

    const [paymentKpiStats] = await sequelize.query(paymentKpiQuery, {
      replacements,
      type: QueryTypes.SELECT,
    });

    console.log("[Dashboard] Policy KPI Stats:", policyKpiStats, "Payment KPI Stats:", paymentKpiStats);

    // ─────────────────────────────────────────────────────────────
    // 5. REMINDERS (Month-wise / Date Filtered)
    // ─────────────────────────────────────────────────────────────
    const reminderReplacements = {
      ...replacements,
      beforeDays,
      afterDays,
      afterDaysNeg: -afterDays,
    };

    let reminderDateFilterSQL = "";
    if (fromDate && toDate) {
      reminderDateFilterSQL = ` AND CAST(r.POLICY_END_DATE AS DATE) BETWEEN :fromDate AND :toDate`;
      reminderReplacements.fromDate = fromDate;
      reminderReplacements.toDate = toDate;
    } else if (fromDate) {
      reminderDateFilterSQL = ` AND CAST(r.POLICY_END_DATE AS DATE) >= :fromDate`;
      reminderReplacements.fromDate = fromDate;
    } else if (toDate) {
      reminderDateFilterSQL = ` AND CAST(r.POLICY_END_DATE AS DATE) <= :toDate`;
      reminderReplacements.toDate = toDate;
    } else {
      reminderDateFilterSQL = ` AND CAST(r.POLICY_END_DATE AS DATE) BETWEEN DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1) AND EOMONTH(GETDATE())`;
    }

    const reminderQuery = `
      ;WITH BlockedVehicles AS (
        SELECT
          REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '') AS VEH_KEY
        FROM dbo.FOLLOWUP_DETAILS f
        JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = f.TRAN_ID
        WHERE f.FOLLOWUP_DATE IS NOT NULL
          AND ISNULL(m.EXPORT_TYPE, 1) = :exportType
          ${locFilterMSQL}
        GROUP BY REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '')
        HAVING
          MAX(CASE
                WHEN UPPER(LTRIM(RTRIM(f.FOLLOWUP_STATUS))) IN ('NOT_INTERESTED','RENEWED')
                THEN 1 ELSE 0
              END) = 1
          OR MAX(CAST(f.FOLLOWUP_DATE AS DATE)) >= CAST(GETDATE() AS DATE)
      )
      SELECT
        COUNT(CASE
          WHEN CAST(r.POLICY_END_DATE AS DATE) >= CAST(GETDATE() AS DATE)
               THEN 1 END) AS beforeExpiryReminders,

        COUNT(CASE
          WHEN CAST(r.POLICY_END_DATE AS DATE) < CAST(GETDATE() AS DATE)
               THEN 1 END) AS afterExpiryReminders

      FROM dbo.INSU_RENEWAL r
      JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = r.TRAN_ID
      LEFT JOIN BlockedVehicles bv
        ON bv.VEH_KEY = REPLACE(UPPER(r.VEHICAL_REG_NO), ' ', '')
      WHERE ISNULL(r.EXPORT_TYPE, 1) = :exportType
        AND ISNULL(m.EXPORT_TYPE, 1) = :exportType
        AND bv.VEH_KEY IS NULL
        ${locFilterMSQL}
        ${reminderDateFilterSQL}
    `;

    const [reminderStats] = await sequelize.query(reminderQuery, {
      replacements: reminderReplacements,
      type: QueryTypes.SELECT,
    });

    console.log("[Dashboard] Reminder Stats:", reminderStats);

    // ─────────────────────────────────────────────────────────────
    // 6. FOLLOWUP BREAKDOWN
    // ─────────────────────────────────────────────────────────────
    const followupBreakdownQuery = `
      ;WITH FU AS (
        SELECT
          CASE
            WHEN NULLIF(LTRIM(RTRIM(f.FOLLOWUP_STATUS)), '') IS NULL THEN 'NO_STATUS'
            ELSE UPPER(LTRIM(RTRIM(f.FOLLOWUP_STATUS)))
          END AS status
        FROM dbo.FOLLOWUP_DETAILS f
        INNER JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = f.TRAN_ID
        OUTER APPLY (
          SELECT TOP 1 ir.POLICY_END_DATE
          FROM dbo.INSU_RENEWAL ir
          WHERE ir.TRAN_ID = m.UTD AND ISNULL(ir.EXPORT_TYPE, 1) = :exportType
          ORDER BY ir.UTD DESC
        ) r
        WHERE ISNULL(m.EXPORT_TYPE, 1) = :exportType
          ${locFilterMSQL}
          ${dateFilterSQL}
      )
      SELECT
        status,
        COUNT(*) AS count
      FROM FU
      GROUP BY status
      ORDER BY count DESC
    `;

    const followupBreakdown = await sequelize.query(followupBreakdownQuery, {
      replacements,
      type: QueryTypes.SELECT,
    });

    // ─────────────────────────────────────────────────────────────
    // 7. PERIOD-WISE TREND
    // ─────────────────────────────────────────────────────────────
    const periodTrendQuery = `
      SELECT
        ${periodFormatExpr}                              AS periodLabel,

        COUNT(DISTINCT r.UTD)                            AS totalExpiring,

        COUNT(CASE
          WHEN p.PYMT_STATUS = 1 AND p.ACNT_APPR_STATUS = 1
          THEN p.UTD END)                            AS approvedCount,

        COUNT(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS IS NULL OR (p.ACNT_APPR_STATUS = 0 AND NULLIF(p.ACNT_APPR_REMARK, '') IS NULL AND NULLIF(p.ACNT_APPR_CODE, '') IS NULL AND p.ACNT_APPR_DATE IS NULL))
          THEN p.UTD END)                            AS pendingCount,

        COUNT(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS = 2 OR (p.ACNT_APPR_STATUS = 0 AND (NULLIF(p.ACNT_APPR_REMARK, '') IS NOT NULL OR NULLIF(p.ACNT_APPR_CODE, '') IS NOT NULL OR p.ACNT_APPR_DATE IS NOT NULL)))
          THEN p.UTD END)                            AS rejectedCount,

        ISNULL(SUM(CASE
          WHEN p.PYMT_STATUS = 1 AND p.ACNT_APPR_STATUS = 1
          THEN p.PYMT_AMOUNT ELSE 0 END), 0)             AS collectedAmount,

        ISNULL(SUM(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS = 2 OR (p.ACNT_APPR_STATUS = 0 AND (NULLIF(p.ACNT_APPR_REMARK, '') IS NOT NULL OR NULLIF(p.ACNT_APPR_CODE, '') IS NOT NULL OR p.ACNT_APPR_DATE IS NOT NULL)))
          THEN p.PYMT_AMOUNT ELSE 0 END), 0)             AS rejectedAmount,

        ISNULL(SUM(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS IS NULL OR (p.ACNT_APPR_STATUS = 0 AND NULLIF(p.ACNT_APPR_REMARK, '') IS NULL AND NULLIF(p.ACNT_APPR_CODE, '') IS NULL AND p.ACNT_APPR_DATE IS NULL))
          THEN p.PYMT_AMOUNT ELSE 0 END), 0)             AS pendingAmount,

        ISNULL(SUM(COALESCE(NULLIF(r.PREMIUM_AMOUNT, 0), NULLIF(p.PREMIUM_AMOUNT, 0), p.PYMT_AMOUNT, 0)), 0) AS totalPremiumAmount,

        COUNT(DISTINCT CASE
          WHEN CAST(r.POLICY_END_DATE AS DATE) >= CAST(GETDATE() AS DATE)
          THEN r.UTD END)                                AS activePoliciesCount,

        COUNT(DISTINCT CASE
          WHEN CAST(r.POLICY_END_DATE AS DATE) < CAST(GETDATE() AS DATE)
          THEN r.UTD END)                                AS expiredPoliciesCount

      FROM dbo.INSU_RENEWAL r
      INNER JOIN dbo.INSU_RENEWAL_MST m
        ON m.UTD = r.TRAN_ID
        AND ISNULL(m.EXPORT_TYPE, 1) = :exportType
        ${locFilterMSQL}
      LEFT JOIN dbo.INSU_RENEWAL_PYMT p
        ON p.TRAN_ID = m.UTD
      WHERE ISNULL(r.EXPORT_TYPE, 1) = :exportType
        AND r.POLICY_END_DATE IS NOT NULL
        ${periodDateFilterSQL}
      GROUP BY ${periodFormatExpr}
      ORDER BY ${periodOrderExpr} ASC
    `;

    const periodTrend = await sequelize.query(periodTrendQuery, {
      replacements: periodReplacements,
      type: QueryTypes.SELECT,
    });

    console.log(
      `[Dashboard] Period Trend: type=${periodTypeFinal}, rows=${periodTrend.length}`
    );

    // ─────────────────────────────────────────────────────────────
    // 7-B. PERIOD-WISE SUMMARY KPI
    // ─────────────────────────────────────────────────────────────
    const periodSummaryQuery = `
      SELECT
        COUNT(DISTINCT r.UTD)                            AS periodTotalPolicies,

        COUNT(CASE
          WHEN p.PYMT_STATUS = 1 AND p.ACNT_APPR_STATUS = 1
          THEN p.UTD END)                                AS periodApprovedCount,

        COUNT(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS IS NULL OR (p.ACNT_APPR_STATUS = 0 AND NULLIF(p.ACNT_APPR_REMARK, '') IS NULL AND NULLIF(p.ACNT_APPR_CODE, '') IS NULL AND p.ACNT_APPR_DATE IS NULL))
          THEN p.UTD END)                                AS periodPendingCount,

        COUNT(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS = 2 OR (p.ACNT_APPR_STATUS = 0 AND (NULLIF(p.ACNT_APPR_REMARK, '') IS NOT NULL OR NULLIF(p.ACNT_APPR_CODE, '') IS NOT NULL OR p.ACNT_APPR_DATE IS NOT NULL)))
          THEN p.UTD END)                                AS periodRejectedCount,

        ISNULL(SUM(CASE
          WHEN p.PYMT_STATUS = 1 AND p.ACNT_APPR_STATUS = 1
          THEN p.PYMT_AMOUNT ELSE 0 END), 0)             AS periodCollectedAmount,

        ISNULL(SUM(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS IS NULL OR (p.ACNT_APPR_STATUS = 0 AND NULLIF(p.ACNT_APPR_REMARK, '') IS NULL AND NULLIF(p.ACNT_APPR_CODE, '') IS NULL AND p.ACNT_APPR_DATE IS NULL))
          THEN p.PYMT_AMOUNT ELSE 0 END), 0)             AS periodPendingAmount,

        ISNULL(SUM(CASE
          WHEN p.PYMT_STATUS = 1 AND (p.ACNT_APPR_STATUS = 2 OR (p.ACNT_APPR_STATUS = 0 AND (NULLIF(p.ACNT_APPR_REMARK, '') IS NOT NULL OR NULLIF(p.ACNT_APPR_CODE, '') IS NOT NULL OR p.ACNT_APPR_DATE IS NOT NULL)))
          THEN p.PYMT_AMOUNT ELSE 0 END), 0)             AS periodRejectedAmount,

        ISNULL(SUM(COALESCE(NULLIF(r.PREMIUM_AMOUNT, 0), NULLIF(p.PREMIUM_AMOUNT, 0), p.PYMT_AMOUNT, 0)), 0) AS periodTotalPremium

      FROM dbo.INSU_RENEWAL r
      INNER JOIN dbo.INSU_RENEWAL_MST m
        ON m.UTD = r.TRAN_ID
        AND ISNULL(m.EXPORT_TYPE, 1) = :exportType
        ${locFilterMSQL}
      LEFT JOIN dbo.INSU_RENEWAL_PYMT p
        ON p.TRAN_ID = m.UTD
      WHERE ISNULL(r.EXPORT_TYPE, 1) = :exportType
        AND r.POLICY_END_DATE IS NOT NULL
        ${periodDateFilterSQL}
    `;

    const [periodSummary] = await sequelize.query(periodSummaryQuery, {
      replacements: periodReplacements,
      type: QueryTypes.SELECT,
    });

    console.log("[Dashboard] Period Summary:", periodSummary);

    // ─────────────────────────────────────────────────────────────
    // 8. TOP COMPANIES - Filtered by current month CREATED_AT and matched with KPI
    // ─────────────────────────────────────────────────────────────
    const insuranceCompanyDistributionQuery = `
      SELECT TOP 5
        ISNULL(NULLIF(LTRIM(RTRIM(r.POLICY_NAME)), ''), 'Other') AS companyName,
        COUNT(DISTINCT m.UTD) AS policyCount,
        ISNULL(SUM(COALESCE(NULLIF(r.PREMIUM_AMOUNT, 0), NULLIF(p.PREMIUM_AMOUNT, 0), p.PYMT_AMOUNT, 0)), 0) AS totalPremium,
        ISNULL(SUM(CASE WHEN p.PYMT_STATUS = 1 AND p.ACNT_APPR_STATUS = 1 THEN p.PYMT_AMOUNT ELSE 0 END), 0) AS collectedAmount
      FROM dbo.INSU_RENEWAL_MST m
      CROSS APPLY (
        SELECT TOP 1
          ir.POLICY_NAME,
          ir.POLICY_END_DATE,
          ir.PREMIUM_AMOUNT
        FROM dbo.INSU_RENEWAL ir
        WHERE ir.TRAN_ID = m.UTD
          AND ISNULL(ir.EXPORT_TYPE, 1) = :exportType
        ORDER BY ir.UTD DESC
      ) r
      LEFT JOIN dbo.INSU_RENEWAL_PYMT p
        ON p.TRAN_ID = m.UTD
        AND p.PYMT_STATUS = 1
      WHERE ISNULL(m.EXPORT_TYPE, 1) = :exportType
        ${locFilterMSQL}
        ${dateFilterSQL}
      GROUP BY ISNULL(NULLIF(LTRIM(RTRIM(r.POLICY_NAME)), ''), 'Other')
      ORDER BY policyCount DESC, totalPremium DESC
    `;

    const companyDistribution = await sequelize.query(
      insuranceCompanyDistributionQuery,
      { replacements, type: QueryTypes.SELECT }
    );

    console.log("[Dashboard] Company Distribution:", companyDistribution);

    // ─────────────────────────────────────────────────────────────
    // 9. DETAILED TABLES / DRILLDOWN DATA
    // ─────────────────────────────────────────────────────────────
    // A) Recent Expiring (Next 30 days)
    const recentExpiringQuery = `
      SELECT TOP 100
        ISNULL(r.UTD, m.UTD)                          AS UTD,
        m.UTD                                         AS TRAN_ID,
        m.VEHICAL_REG_NO,
        ISNULL(r.CUST_NAME, '')                       AS CUST_NAME,
        CAST(ISNULL(r.CUST_MOB_NO, '') AS NVARCHAR(20)) AS CUST_MOB_NO,
        ISNULL(r.POLICY_NAME, '')                     AS POLICY_NAME,
        CAST(ISNULL(r.POLICY_NUMBER, '') AS NVARCHAR(100)) AS POLICY_NUMBER,
        ISNULL(r.MODEL_NAME, '')                      AS MODEL_NAME,
        ISNULL(NULLIF(r.PREMIUM_AMOUNT, 0), ISNULL(p.PREMIUM_AMOUNT, ISNULL(p.PYMT_AMOUNT, 0))) AS PREMIUM_AMOUNT,
        CONVERT(varchar(10), r.POLICY_START_DATE, 105) AS POLICY_START_DATE,
        CONVERT(varchar(10), r.POLICY_END_DATE, 105)  AS POLICY_END_DATE,
        DATEDIFF(day, CAST(GETDATE() AS DATE), r.POLICY_END_DATE) AS DAYS_TO_EXPIRY,
        'ACTIVE'                                      AS POLICY_STATUS,
        (
          SELECT TOP 1 fd.FOLLOWUP_STATUS
          FROM dbo.FOLLOWUP_DETAILS fd
          WHERE fd.TRAN_ID = m.UTD
          ORDER BY CAST(fd.FOLLOWUP_DATE AS DATE) DESC, fd.UTD DESC
        ) AS LAST_FOLLOWUP_STATUS,
        (
          SELECT TOP 1 CONVERT(varchar(10), fd.FOLLOWUP_DATE, 105)
          FROM dbo.FOLLOWUP_DETAILS fd
          WHERE fd.TRAN_ID = m.UTD
          ORDER BY CAST(fd.FOLLOWUP_DATE AS DATE) DESC, fd.UTD DESC
        ) AS LAST_FOLLOWUP_DATE
      FROM dbo.INSU_RENEWAL_MST m
      OUTER APPLY (
        SELECT TOP 1
          ir.UTD,
          ir.CUST_NAME,
          ir.CUST_MOB_NO,
          ir.POLICY_NAME,
          ir.POLICY_NUMBER,
          ir.MODEL_NAME,
          ir.PREMIUM_AMOUNT,
          ir.POLICY_START_DATE,
          ir.POLICY_END_DATE
        FROM dbo.INSU_RENEWAL ir
        WHERE ir.TRAN_ID = m.UTD
          AND ISNULL(ir.EXPORT_TYPE, 1) = :exportType
        ORDER BY ir.UTD DESC
      ) r
      OUTER APPLY (
        SELECT TOP 1
          pymt.PREMIUM_AMOUNT,
          pymt.PYMT_AMOUNT
        FROM dbo.INSU_RENEWAL_PYMT pymt
        WHERE pymt.TRAN_ID = m.UTD AND pymt.PYMT_STATUS = 1
        ORDER BY pymt.UTD DESC
      ) p
      WHERE ISNULL(m.EXPORT_TYPE, 1) = :exportType
        AND r.POLICY_END_DATE IS NOT NULL
        AND CAST(r.POLICY_END_DATE AS DATE) >= CAST(GETDATE() AS DATE)
        AND CAST(r.POLICY_END_DATE AS DATE) <= DATEADD(day, 30, CAST(GETDATE() AS DATE))
        ${locFilterMSQL}
        ${dateFilterSQL}
      ORDER BY r.POLICY_END_DATE ASC
    `;

    const recentExpiring = await sequelize.query(recentExpiringQuery, {
      replacements,
      type: QueryTypes.SELECT,
    });

    // B) All Policies (Total, Active, Expired)
    // Querying exactly from INSU_RENEWAL_MST m with OUTER APPLY so count is 1:1 identical to KPI Stats
    const allPoliciesQuery = `
      SELECT TOP 500
        ISNULL(r.UTD, m.UTD)                          AS UTD,
        m.UTD                                         AS TRAN_ID,
        m.VEHICAL_REG_NO,
        ISNULL(r.CUST_NAME, '')                       AS CUST_NAME,
        CAST(ISNULL(r.CUST_MOB_NO, '') AS NVARCHAR(20)) AS CUST_MOB_NO,
        ISNULL(r.POLICY_NAME, '')                     AS POLICY_NAME,
        CAST(ISNULL(r.POLICY_NUMBER, '') AS NVARCHAR(100)) AS POLICY_NUMBER,
        ISNULL(r.MODEL_NAME, '')                      AS MODEL_NAME,
        ISNULL(r.DSC_EMPCODE, '')                     AS DSC_EMPCODE,
        ISNULL(r.DSC_NAME, '')                        AS DSC_NAME,
        ISNULL(r.DSC_MOB_NO, '')                      AS DSC_MOB_NO,
        ISNULL(NULLIF(r.PREMIUM_AMOUNT, 0), ISNULL(p.PREMIUM_AMOUNT, ISNULL(p.PYMT_AMOUNT, 0))) AS PREMIUM_AMOUNT,
        CONVERT(varchar(10), r.POLICY_START_DATE, 105) AS POLICY_START_DATE,
        CONVERT(varchar(10), r.POLICY_END_DATE, 105)  AS POLICY_END_DATE,
        CONVERT(varchar(10), ISNULL(r.CREATED_AT, m.CREATED_AT), 105) AS CREATED_AT,
        DATEDIFF(day, CAST(GETDATE() AS DATE), r.POLICY_END_DATE) AS DAYS_TO_EXPIRY,
        CASE
          WHEN r.POLICY_END_DATE IS NULL THEN 'EXPIRED'
          WHEN CAST(r.POLICY_END_DATE AS DATE) >= CAST(GETDATE() AS DATE) THEN 'ACTIVE'
          ELSE 'EXPIRED'
        END AS POLICY_STATUS,
        (
          SELECT TOP 1 fd.FOLLOWUP_STATUS
          FROM dbo.FOLLOWUP_DETAILS fd
          WHERE fd.TRAN_ID = m.UTD
          ORDER BY CAST(fd.FOLLOWUP_DATE AS DATE) DESC, fd.UTD DESC
        ) AS LAST_FOLLOWUP_STATUS,
        (
          SELECT TOP 1 CONVERT(varchar(10), fd.FOLLOWUP_DATE, 105)
          FROM dbo.FOLLOWUP_DETAILS fd
          WHERE fd.TRAN_ID = m.UTD
          ORDER BY CAST(fd.FOLLOWUP_DATE AS DATE) DESC, fd.UTD DESC
        ) AS LAST_FOLLOWUP_DATE
      FROM dbo.INSU_RENEWAL_MST m
      OUTER APPLY (
        SELECT TOP 1
          ir.UTD,
          ir.CUST_NAME,
          ir.CUST_MOB_NO,
          ir.POLICY_NAME,
          ir.POLICY_NUMBER,
          ir.MODEL_NAME,
          ir.PREMIUM_AMOUNT,
          ir.POLICY_START_DATE,
          ir.POLICY_END_DATE,
          ir.DSC_EMPCODE,
          ir.DSC_NAME,
          ir.DSC_MOB_NO,
          ir.CREATED_AT
        FROM dbo.INSU_RENEWAL ir
        WHERE ir.TRAN_ID = m.UTD
          AND ISNULL(ir.EXPORT_TYPE, 1) = :exportType
        ORDER BY ir.UTD DESC
      ) r
      OUTER APPLY (
        SELECT TOP 1
          pymt.PREMIUM_AMOUNT,
          pymt.PYMT_AMOUNT
        FROM dbo.INSU_RENEWAL_PYMT pymt
        WHERE pymt.TRAN_ID = m.UTD AND pymt.PYMT_STATUS = 1
        ORDER BY pymt.UTD DESC
      ) p
      WHERE ISNULL(m.EXPORT_TYPE, 1) = :exportType
        ${locFilterMSQL}
        ${dateFilterSQL}
      ORDER BY r.UTD DESC, m.UTD DESC
    `;

    const allPolicies = await sequelize.query(allPoliciesQuery, {
      replacements,
      type: QueryTypes.SELECT,
    });

    const activePolicies = allPolicies.filter((p) => p.POLICY_STATUS === "ACTIVE");
    const expiredPolicies = allPolicies.filter((p) => p.POLICY_STATUS === "EXPIRED");
    const expiringIn30Days = allPolicies.filter((p) => {
      const d = parseInt(p.DAYS_TO_EXPIRY, 10);
      return Number.isFinite(d) && d >= 0 && d <= 30;
    });
    const premiumPolicies = allPolicies.filter((p) => {
      const prem = parseFloat(p.PREMIUM_AMOUNT || 0);
      return Number.isFinite(prem) && prem > 0;
    });

    const allPaymentsQuery = `
      ;WITH DistinctPymt AS (
        SELECT 
          p.UTD,
          p.TRAN_ID,
          p.PYMT_AMOUNT,
          p.PYMT_MODE,
          p.PYMT_DATE,
          p.BANK_NAME,
          p.ACNT_APPR_STATUS,
          p.ACNT_APPR_REMARK,
          p.ACNT_APPR_CODE,
          p.ACNT_APPR_DATE,
          p.REMARKS,
          p.PYMT_REMARK,
          ROW_NUMBER() OVER (
            PARTITION BY COALESCE(NULLIF(LTRIM(RTRIM(m.VEHICAL_REG_NO)), ''), CAST(p.TRAN_ID AS VARCHAR))
            ORDER BY p.UTD DESC
          ) AS rn
        FROM dbo.INSU_RENEWAL_PYMT p
        LEFT JOIN dbo.INSU_RENEWAL_MST m
          ON m.UTD = p.TRAN_ID
        WHERE p.PYMT_STATUS = 1
          ${locFilterPSQL}
          ${pymtDateFilterSQL}
      )
      SELECT TOP 500
        p.UTD                                         AS UTD,
        p.UTD                                         AS PYMT_UTD,
        p.TRAN_ID,
        COALESCE(m.VEHICAL_REG_NO, r.VEHICAL_REG_NO, '') AS VEHICAL_REG_NO,
        ISNULL(r.CUST_NAME, '')                       AS CUST_NAME,
        CAST(ISNULL(r.CUST_MOB_NO, '') AS NVARCHAR(20)) AS CUST_MOB_NO,
        ISNULL(r.POLICY_NAME, '')                     AS POLICY_NAME,
        CAST(ISNULL(r.POLICY_NUMBER, '') AS NVARCHAR(100)) AS POLICY_NUMBER,
        ISNULL(r.MODEL_NAME, '')                      AS MODEL_NAME,
        ISNULL(p.PYMT_AMOUNT, 0)                      AS PYMT_AMOUNT,
        p.PYMT_MODE,
        CONVERT(varchar(10), p.PYMT_DATE, 105)        AS PYMT_DATE,
        p.BANK_NAME,
        p.ACNT_APPR_STATUS,
        CASE 
          WHEN p.ACNT_APPR_STATUS = 1 THEN 'APPROVED'
          WHEN p.ACNT_APPR_STATUS = 2 OR (p.ACNT_APPR_STATUS = 0 AND (NULLIF(p.ACNT_APPR_REMARK, '') IS NOT NULL OR NULLIF(p.ACNT_APPR_CODE, '') IS NOT NULL OR p.ACNT_APPR_DATE IS NOT NULL)) THEN 'REJECTED'
          ELSE 'PENDING'
        END AS APPROVAL_STATUS_LABEL,
        p.ACNT_APPR_REMARK,
        p.REMARKS,
        p.PYMT_REMARK
      FROM DistinctPymt p
      LEFT JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = p.TRAN_ID
      OUTER APPLY (
        SELECT TOP 1 ir.CUST_NAME, ir.CUST_MOB_NO, ir.POLICY_NAME, ir.POLICY_NUMBER, ir.MODEL_NAME, ir.POLICY_END_DATE, ir.VEHICAL_REG_NO
        FROM dbo.INSU_RENEWAL ir
        WHERE (ir.TRAN_ID = m.UTD OR (m.VEHICAL_REG_NO IS NOT NULL AND REPLACE(UPPER(ir.VEHICAL_REG_NO), ' ', '') = REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '')) OR ir.TRAN_ID = p.TRAN_ID)
          AND ISNULL(ir.EXPORT_TYPE, 1) = :exportType
        ORDER BY ir.UTD DESC
      ) r
      WHERE p.rn = 1
      ORDER BY p.UTD DESC
    `;

    const allPayments = await sequelize.query(allPaymentsQuery, {
      replacements,
      type: QueryTypes.SELECT,
    });

    const approvedPayments = allPayments.filter((p) => p.APPROVAL_STATUS_LABEL === "APPROVED");
    const pendingPayments = allPayments.filter((p) => p.APPROVAL_STATUS_LABEL === "PENDING");
    const rejectedPayments = allPayments.filter((p) => p.APPROVAL_STATUS_LABEL === "REJECTED");

    // D) Reminders List (Before & After Expiry)
    const allRemindersQuery = `
      ;WITH BlockedVehicles AS (
        SELECT
          REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '') AS VEH_KEY
        FROM dbo.FOLLOWUP_DETAILS f
        JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = f.TRAN_ID
        WHERE f.FOLLOWUP_DATE IS NOT NULL
          AND ISNULL(m.EXPORT_TYPE, 1) = :exportType
          ${locFilterMSQL}
        GROUP BY REPLACE(UPPER(m.VEHICAL_REG_NO), ' ', '')
        HAVING
          MAX(CASE
                WHEN UPPER(LTRIM(RTRIM(f.FOLLOWUP_STATUS))) IN ('NOT_INTERESTED','RENEWED')
                THEN 1 ELSE 0
              END) = 1
          OR MAX(CAST(f.FOLLOWUP_DATE AS DATE)) >= CAST(GETDATE() AS DATE)
      )
      SELECT TOP 500
        r.UTD,
        r.TRAN_ID,
        m.VEHICAL_REG_NO,
        r.CUST_NAME,
        CAST(r.CUST_MOB_NO AS NVARCHAR(20))          AS CUST_MOB_NO,
        r.POLICY_NAME,
        CAST(r.POLICY_NUMBER AS NVARCHAR(100))        AS POLICY_NUMBER,
        r.MODEL_NAME,
        CONVERT(varchar(10), r.POLICY_END_DATE, 105)  AS POLICY_END_DATE,
        DATEDIFF(day, CAST(GETDATE() AS DATE), r.POLICY_END_DATE) AS DAYS_TO_EXPIRY,
        CASE
          WHEN CAST(r.POLICY_END_DATE AS DATE) >= CAST(GETDATE() AS DATE) THEN 'BEFORE_EXPIRY'
          ELSE 'AFTER_EXPIRY'
        END AS REMINDER_TYPE,
        (
          SELECT TOP 1 fd.FOLLOWUP_STATUS
          FROM dbo.FOLLOWUP_DETAILS fd
          WHERE fd.TRAN_ID = m.UTD
          ORDER BY CAST(fd.FOLLOWUP_DATE AS DATE) DESC, fd.UTD DESC
        ) AS LAST_FOLLOWUP_STATUS,
        (
          SELECT TOP 1 CONVERT(varchar(10), fd.FOLLOWUP_DATE, 105)
          FROM dbo.FOLLOWUP_DETAILS fd
          WHERE fd.TRAN_ID = m.UTD
          ORDER BY CAST(fd.FOLLOWUP_DATE AS DATE) DESC, fd.UTD DESC
        ) AS LAST_FOLLOWUP_DATE
      FROM dbo.INSU_RENEWAL r
      JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = r.TRAN_ID
      LEFT JOIN BlockedVehicles bv ON bv.VEH_KEY = REPLACE(UPPER(r.VEHICAL_REG_NO), ' ', '')
      WHERE ISNULL(r.EXPORT_TYPE, 1) = :exportType
        AND ISNULL(m.EXPORT_TYPE, 1) = :exportType
        AND bv.VEH_KEY IS NULL
        ${locFilterMSQL}
        ${reminderDateFilterSQL}
      ORDER BY r.POLICY_END_DATE ASC
    `;

    const allReminders = await sequelize.query(allRemindersQuery, {
      replacements: reminderReplacements,
      type: QueryTypes.SELECT,
    });

    const beforeExpiryReminders = allReminders.filter((r) => r.REMINDER_TYPE === "BEFORE_EXPIRY");
    const afterExpiryReminders = allReminders.filter((r) => r.REMINDER_TYPE === "AFTER_EXPIRY");

    // ─────────────────────────────────────────────────────────────
    // RESPONSE
    // ─────────────────────────────────────────────────────────────
    return res.status(200).json({
      success: true,
      Message: "Dashboard metrics fetched successfully.",
      data: {
        config: {
          beforeDays,
          afterDays,
          exportType: exportTypeFinal,
          locCodes,
          periodType: periodTypeFinal,
          periodLabel,
          dateFilter: fromDate && toDate ? { fromDate, toDate } : null,
          _debug: {
            periodRangeStart,
            periodRangeEnd,
            periodTrendRows: periodTrend.length,
            dateFilterSQL,
            periodDateFilterSQL,
          },
        },
        summary: {
          totalPolicies: parseInt(policyKpiStats?.totalPolicies ?? 0, 10),
          activePolicies: parseInt(policyKpiStats?.activePolicies ?? 0, 10),
          expiredPolicies: parseInt(policyKpiStats?.expiredPolicies ?? 0, 10),
          expiringIn30Days: parseInt(policyKpiStats?.expiringIn30Days ?? 0, 10),
          reminders: {
            beforeExpiry: parseInt(
              reminderStats?.beforeExpiryReminders ?? 0,
              10
            ),
            afterExpiry: parseInt(reminderStats?.afterExpiryReminders ?? 0, 10),
            totalDue:
              parseInt(reminderStats?.beforeExpiryReminders ?? 0, 10) +
              parseInt(reminderStats?.afterExpiryReminders ?? 0, 10),
            beforeDays,
            afterDays,
          },
          payments: {
            approvedCount: parseInt(paymentKpiStats?.approvedPaymentsCount ?? 0, 10),
            rejectedCount: parseInt(paymentKpiStats?.rejectedPaymentsCount ?? 0, 10),
            pendingCount: parseInt(paymentKpiStats?.pendingPaymentsCount ?? 0, 10),
            totalCount: parseInt(paymentKpiStats?.totalPaymentsCount ?? 0, 10),
            totalCollectedAmount: parseFloat(
              paymentKpiStats?.totalCollectedAmount ?? 0
            ),
            totalRejectedAmount: parseFloat(
              paymentKpiStats?.totalRejectedAmount ?? 0
            ),
            totalPendingAmount: parseFloat(paymentKpiStats?.totalPendingAmount ?? 0),
            totalPremiumAmount: parseFloat(
              (policyKpiStats?.totalPremiumAmount || paymentKpiStats?.totalPaymentPremium) ?? 0
            ),
          },
          periodSummary: {
            periodType: periodTypeFinal,
            periodLabel,
            totalPolicies: parseInt(
              periodSummary?.periodTotalPolicies ?? 0,
              10
            ),
            collectedAmount: parseFloat(
              periodSummary?.periodCollectedAmount ?? 0
            ),
            pendingAmount: parseFloat(periodSummary?.periodPendingAmount ?? 0),
            rejectedAmount: parseFloat(
              periodSummary?.periodRejectedAmount ?? 0
            ),
            totalPremium: parseFloat(periodSummary?.periodTotalPremium ?? 0),
            approvedCount: parseInt(
              periodSummary?.periodApprovedCount ?? 0,
              10
            ),
            rejectedCount: parseInt(
              periodSummary?.periodRejectedCount ?? 0,
              10
            ),
            pendingCount: parseInt(periodSummary?.periodPendingCount ?? 0, 10),
          },
        },
        charts: {
          followupBreakdown,
          periodTrend,
          monthlyTrend: periodTypeFinal === "monthly" ? periodTrend : [],
          companyDistribution,
        },
        tables: {
          recentExpiring,
          totalPolicies: allPolicies,
          activePolicies,
          expiredPolicies,
          expiringIn30Days,
          premiumPolicies,
          approvedPayments,
          pendingPayments,
          rejectedPayments,
          totalPayments: allPayments,
          beforeExpiryReminders,
          afterExpiryReminders,
          totalReminders: allReminders,
        },
      },
    });
  } catch (error) {
    console.error("Error in getInsuranceDashboardMetrics:", error);
    return res.status(500).json({
      success: false,
      Message: "An error occurred while fetching dashboard metrics.",
      Error: error.message,
    });
  } finally {
    await sequelize.close();
  }
};

exports.OD_report_With_Pic = async function (req, res) {
  const sequelize = await dbname(req, req.query.compcode);
  try {
    const data = req.body;
    const dateto = data.dateto;
    const dateFrom = data.dateFrom;
    const region = data.region;
    const channel = data.channel;
    const Cluster = data.Cluster;
    const location = data.location;
    const section = data.section;
    const department = data.department;
    const designation = data.designation;
    const emptype = data.emptype;
    const empcode = data.empcode;
    const reporttype = data.reporttype;
    const LeftEmp = data.LeftEmp;

    // Pagination parameters
    const page = parseInt(data.page) || 1;
    const pageSize = parseInt(data.pageSize) || 10;
    const offset = (page - 1) * pageSize;

    var query;

    query = `select empcode from employeemaster where export_type<3 `;

    if (region) query += ` and sal_region in (${region}) `;
    if (channel) query += ` and channel in (${channel}) `;
    if (Cluster) query += ` and CLUSTER in (${Cluster}) `;
    const user_location = data.user_location;
    if (parseInt(user_location)) {
      if (location) query += `and location in (${location}) `;
    } else {
      if (location)
        query += `and location in (select misc_code from misc_mst where misc_type = 85 and misc_hod in (${location})) `;
    }
    if (section) query += `and section in (${section}) `;
    if (department) query += `and division in (${department}) `;
    if (designation) {
      const desigList = designation
        .split(",")
        .map((d) => `''${d.trim()}''`)
        .join(",");
      query += ` AND EMPLOYEEDESIGNATION IN (${desigList})`;
    }

    if (emptype) query += `and emptype in (${emptype}) `;
    if (empcode) {
      query += `and empcode in (''${empcode}'') `;
    }
    if (LeftEmp != 1) {
      query += ` AND lastwor_date IS NULL `;
    }
    console.log(query, "this is query ");

    let moduleCode = "attdence"; // default

    const moduleCodeQuery = `SELECT mp_keyword FROM COMP_KEYDATA`;
    const [keyData] = await sequelize.query(moduleCodeQuery);

    if (keyData.length > 0 && keyData[0].mp_keyword) {
      moduleCode = keyData[0].mp_keyword;
    }

    // Total count query
    const countQuery = `
SELECT COUNT(DISTINCT atn.emp_code) AS total
FROM attendancetable atn
LEFT JOIN employeemaster emp ON emp.empcode = atn.emp_code
WHERE
    atn.dateoffice BETWEEN '${dateFrom}' AND '${dateto}'
    AND atn.mipunch_reason NOT IN (4,9,5,6,10,104,120,121,122,123)
    AND atn.emp_code IN (${query})
`;

    const [countResult] = await sequelize.query(countQuery);
    const totalRecords = countResult[0]?.total || 0;
    const totalPages = Math.ceil(totalRecords / pageSize);

    const main_query2Copy = `
WITH RankedData AS (
SELECT
    MAX(emp.empcode) AS [Employee Code],
    MAX(CONCAT(emp.title, ' ', emp.empfirstname, ' ', emp.emplastname)) AS [Employee Name],
    MAX(emp.currentjoindate) AS [Date Of Joining],
    MAX(emp.MOBILE_NO) AS [Mobile No.],
    (SELECT TOP 1 comp_name FROM comp_mst) AS Company,
    MAX(ch.misc_name) AS Channel,
    MAX(cl.misc_name) AS Cluster,
    MAX(br.misc_name) AS Branch,
    MAX(sec.misc_name) AS Section,
    MAX(dept.misc_name) AS Department,
    MAX(reg.misc_name) AS Region,
    MAX(emp.employeedesignation) AS Designation,

FORMAT(MIN(atn.mp_in1), 'yyyy-MM-dd HH:mm:ss.fff''Z''') AS [Start From],
FORMAT(MAX(atn.mp_out1), 'yyyy-MM-dd HH:mm:ss.fff''Z''') AS [End From],

    MAX(std.misc_name) AS [OD Type],
    MAX(atn.shiftstarttime) AS [Shift Start Time],
    MAX(atn.shiftendtime) AS [Shift End Time],
    MAX(atn.status) AS [ATTN. STATUS],
FORMAT(MAX(atn.Mispunch_applied_on), 'yyyy-MM-dd HH:mm:ss.fff''Z''') AS [Mispunch IN Applied On],
FORMAT(MAX(atn.Mispunch_out_applied_on), 'yyyy-MM-dd HH:mm:ss.fff''Z''') AS [Mispunch OUT Applied On],
MAX(atn.Mis_Enterby) AS [Mispunch Enter By],

       CASE
         WHEN (MAX(atn.mp_in1) IS NOT NULL OR MAX(atn.mp_in1)  = '')
     AND (MAX(atn.mp_out1) IS NOT NULL OR MAX(atn.mp_out1) = '')
        THEN 'Mispunch In/Out'
    WHEN (MAX(atn.mp_in1)  IS NULL OR MAX(atn.mp_in1)  = '')
     AND (MAX(atn.mp_out1) IS NULL OR MAX(atn.mp_out1) = '')
        THEN 'Missing In/Out'

    WHEN MAX(atn.mp_out1) IS NOT NULL AND MAX(atn.mp_out1) <> ''
        THEN 'Mispunch Out'

    WHEN MAX(atn.mp_in1) IS NOT NULL AND MAX(atn.mp_in1) <> ''
        THEN 'Mispunch In'

    ELSE NULL
END AS [Applied Type],

    SUM(atn.Short_Lev) AS [Total_OD_Days],

    CASE
    WHEN (MAX(atn.Appr_3_Stat) = 1)
          THEN 'Final Acceptance'

        WHEN MAX(atn.Appr_1_Stat) = 0
          OR MAX(atn.Appr_2_Stat) = 0
          OR MAX(atn.Appr_3_Stat) = 0 THEN 'Rejected'

        WHEN MAX(atn.Appr_1_Stat) = 1
          AND (MAX(atn.Appr_2_Stat) = 1 OR MAX(atn.Appr_2_Code) IS NULL)
          AND (MAX(atn.Appr_3_Stat) = 1 OR MAX(atn.Appr_3_Code) IS NULL)
          THEN 'Final Acceptance'

        ELSE 'Pending'
    END AS [OD Status],
MAX(mirmk.Misc_name) AS [Employee  Remark],
    MAX(atn.SPL_REMARK) AS [Employee SPL Remark],

    -- Approver 1
CASE
  WHEN MAX(atn.Appr_1_Code) IS NULL
    THEN MAX(am.approver1_A)
  ELSE MAX(atn.Appr_1_Code)
END AS [Approver 1 Emp code],

CASE
  WHEN MAX(atn.Appr_1_Code) IS NULL
    THEN MAX(CONCAT(M1.empfirstname,' ',M1.emplastname))
  ELSE MAX(CONCAT(Appr1.empfirstname,' ',Appr1.emplastname))
END AS [Approver 1 Name],

 CASE
        WHEN MAX(atn.Appr_1_Stat) = 1 THEN 'Final Acceptance'
        WHEN MAX(atn.Appr_1_Stat) = 0 THEN 'Rejected'
        WHEN MAX(atn.Appr_1_Stat) IS NULL THEN 'Pending'
        ELSE ''
    END AS [Approver 1 Action Status],
    MAX(atn.Appr_1_Rem) AS [Approver 1 Remarks],
    MAX(atn.Appr_1_date) AS [Approver 1 Action Taken on],

    -- Approver 2
     CASE
  WHEN MAX(atn.Appr_2_Code) IS NULL
    THEN MAX(am.approver2_A)
  ELSE MAX(atn.Appr_2_Code)
END AS [Approver 2 Emp code],
CASE
  WHEN MAX(atn.Appr_2_Code) IS NULL
    THEN MAX(CONCAT(M2.empfirstname,' ',M2.emplastname))
  ELSE MAX(CONCAT(Appr2.empfirstname,' ',Appr2.emplastname))
END AS [Approver 2 Name],

CASE
        WHEN MAX(atn.Appr_1_Stat) IN (0,1) THEN
            CASE
                WHEN MAX(atn.Appr_2_Stat) = 1 THEN 'Final Acceptance'
                WHEN MAX(atn.Appr_2_Stat) = 0 THEN 'Rejected'
                ELSE ''
            END
        ELSE 'Pending'
    END AS [Approver 2 Action Status],
    MAX(atn.Appr_2_Rem) AS [Approver 2 Remarks],
    MAX(atn.Appr_2_date) AS [Approver 2 Action Taken on],

    -- Approver 3
   CASE
  WHEN MAX(atn.Appr_3_Code) IS NULL
    THEN MAX(am.approver3_A)
  ELSE MAX(atn.Appr_3_Code)
END AS [Approver 3 Emp code],

CASE
  WHEN MAX(atn.Appr_3_Code) IS NULL
    THEN MAX(CONCAT(M3.empfirstname,' ',M3.emplastname))
  ELSE MAX(CONCAT(Appr3.empfirstname,' ',Appr3.emplastname))
END AS [Approver 3 Name] ,
CASE
    WHEN MAX(atn.Appr_3_Stat) = 1 THEN 'Final Acceptance'
    WHEN MAX(atn.Appr_3_Stat) = 0 THEN 'Rejected'
    WHEN MAX(atn.Appr_3_Stat) IS NULL THEN ''
    ELSE ''
END AS [Approver 3 Action Status],
    MAX(atn.Appr_3_Rem) AS [Approver 3 Remarks],
    MAX(atn.Appr_3_date) AS [Approver 3 Action Taken on],

    -- ✅ Mispunch In/Out Images - EXACT tumhari di hui query pattern se
    (select top 1
        'https://erp.autovyn.com/backend/fetch?filePath=' + DOC_PATH
     from EMP_DOCS
     where EMP_CODE collate database_default = atn.emp_code collate database_default  
       and misspunch_inout = 1
       and dateoffice = atn.dateoffice
     order by utd desc) as fileIn,

    (select top 1
        'https://erp.autovyn.com/backend/fetch?filePath=' + DOC_PATH
     from EMP_DOCS
     where EMP_CODE collate database_default = atn.emp_code collate database_default  
       and misspunch_inout = 2
       and dateoffice = atn.dateoffice
     order by utd desc) as fileOut,

    ROW_NUMBER() OVER (ORDER BY MAX(emp.empcode), atn.Mispunch_applied_on, atn.Mispunch_out_applied_on) AS RowNum

FROM attendancetable atn
LEFT JOIN employeemaster emp ON emp.empcode = atn.emp_code

LEFT JOIN employeemaster Appr1 ON Appr1.empcode = atn.Appr_1_Code
LEFT JOIN employeemaster Appr2 ON Appr2.empcode = atn.Appr_2_Code
LEFT JOIN employeemaster Appr3 ON Appr3.empcode = atn.Appr_3_Code
LEFT JOIN misc_mst reg ON reg.misc_type = 91 AND reg.misc_code = emp.sal_region
LEFT JOIN misc_mst ch ON ch.misc_type = 627 AND ch.misc_code = emp.channel
LEFT JOIN misc_mst cl ON cl.misc_type = 626 AND cl.misc_code = emp.CLUSTER
LEFT JOIN misc_mst br ON br.misc_type = 85 AND br.misc_code = emp.location
LEFT JOIN misc_mst sec ON sec.misc_type = 81 AND sec.misc_code = emp.section
LEFT JOIN misc_mst dept ON dept.misc_type = 68 AND dept.misc_code = emp.division
LEFT JOIN misc_mst std ON std.misc_type = 92 AND std.misc_code = atn.mipunch_reason
LEFT JOIN misc_mst mirmk ON mirmk.misc_type = 93 AND mirmk.misc_code = atn.MI_Remark
LEFT JOIN Approval_Matrix am
  ON am.module_code = '${moduleCode}'
 AND am.EmpCode = emp.empcode

LEFT JOIN employeemaster M1 ON M1.empcode = am.approver1_A
LEFT JOIN employeemaster M2 ON M2.empcode = am.approver2_A
LEFT JOIN employeemaster M3 ON M3.empcode = am.approver3_A

WHERE
    atn.dateoffice BETWEEN '${dateFrom}' AND '${dateto}'
    AND atn.mipunch_reason NOT IN (4,9,5,6,10,104,120,121,122,123)
    AND atn.emp_code IN (${query})

GROUP BY
    atn.Mispunch_applied_on,
    atn.Mispunch_out_applied_on,
    atn.mipunch_reason,
    atn.emp_code,
    atn.dateoffice
)

SELECT * FROM RankedData
WHERE RowNum > ${offset} AND RowNum <= ${offset + pageSize}
ORDER BY RowNum
`;

    let txnDetails;

    console.log(main_query2Copy, "main_query2Copy");
    txnDetails = await sequelize.query(main_query2Copy);

    return res.status(200).json({
      success: true,
      data: txnDetails[0],
      pagination: {
        currentPage: page,
        pageSize: pageSize,
        totalRecords: totalRecords,
        totalPages: totalPages,
        hasNextPage: page < totalPages,
        hasPreviousPage: page > 1,
      },
    });
  } catch (e) {
    console.log("FULL ERROR:", e);
    console.log("ORIGINAL DB ERROR:", e.original);
    const realErrorMessage =
      e.original?.message ||
      e.parent?.message ||
      e.message ||
      "Unknown error occurred";

    return res.status(500).json({
      success: false,
      message: "Internal Server Error",
      error: realErrorMessage,
    });
  } finally {
    await sequelize.close();
  }
};

exports.attendance_pivot_report = async function (req, res) {
  let sequelize;

  const toYMD = (val) => {
    if (!val) return null;
    if (typeof val === "string") return val.slice(0, 10);
    return new Date(val).toISOString().slice(0, 10);
  };

  const safeCsvNumbers = (v) => {
    if (!v) return "";
    return String(v)
      .split(",")
      .map((x) => parseInt(String(x).trim(), 10))
      .filter((n) => Number.isFinite(n))
      .join(",");
  };

  try {
    const data = req.body;

    const dateto = data.dateto;
    const dateFrom = data.dateFrom;

    // ✅ page/limit must be let (so we can override if needed)
    let page = parseInt(data.page, 10) || 1;
    let limit = parseInt(data.limit, 10) || 10;

    if (!dateFrom || !dateto) {
      return res.status(400).json({
        success: false,
        message: "Date range is required",
      });
    }

    sequelize = await dbname(req, req.query.compcode);

    const startDate = toYMD(dateFrom);
    const endDate = toYMD(dateto);

    const region = safeCsvNumbers(data.region);
    const channel = safeCsvNumbers(data.channel);
    const Cluster = safeCsvNumbers(data.Cluster);
    const location = safeCsvNumbers(data.location);
    const section = safeCsvNumbers(data.section);
    const department = safeCsvNumbers(data.department);
    const emptype = safeCsvNumbers(data.emptype);

    // ✅ GET loc_code FROM HEADER
    const loc_code = safeCsvNumbers(req.headers.loc_code);

    const designation = data.designation;
    const empcode = data.empcode ? String(data.empcode).trim() : null;
    const LeftEmp = data.LeftEmp;

    // ✅ Build filter conditions
    let filterConditions = `em.export_type < 3`;

    if (region) filterConditions += ` AND em.sal_region IN (${region})`;
    if (channel) filterConditions += ` AND em.channel IN (${channel})`;
    if (Cluster) filterConditions += ` AND em.CLUSTER IN (${Cluster})`;

    const user_location = data.user_location;

    // ✅ LOCATION FILTER - Check both location parameter and loc_code header
    if (parseInt(user_location, 10)) {
      // Direct location codes
      if (location) {
        filterConditions += ` AND em.location IN (${location})`;
      }
      // Or loc_code from header
      else if (loc_code) {
        filterConditions += ` AND em.location IN (${loc_code})`;
      }
    } else {
      // HOD-based location filter
      if (location) {
        filterConditions += ` AND em.location IN (
            SELECT misc_code FROM misc_mst WHERE misc_type = 85 AND misc_hod IN (${location})
          )`;
      }
      // Or loc_code header (HOD-based)
      else if (loc_code) {
        filterConditions += ` AND em.location IN (
            SELECT misc_code FROM misc_mst WHERE misc_type = 85 AND misc_hod IN (${loc_code})
          )`;
      }
    }

    if (section) filterConditions += ` AND em.section IN (${section})`;
    if (department) filterConditions += ` AND em.division IN (${department})`;

    if (designation) {
      const desigList = String(designation)
        .split(",")
        .map((d) => `'${d.trim().replace(/'/g, "''")}'`)
        .join(",");
      filterConditions += ` AND em.EMPLOYEEDESIGNATION IN (${desigList})`;
    }

    if (emptype) filterConditions += ` AND em.emptype IN (${emptype})`;

    if (empcode) {
      filterConditions += ` AND LTRIM(RTRIM(em.empcode)) IN ('${empcode.replace(/'/g, "''")}')`;
      // ✅ IMPORTANT: for single employee, return all records
      page = 1;
      limit = 50000;
    }

    if (LeftEmp != 1) {
      filterConditions += ` AND em.lastwor_date IS NULL`;
    }

    // ✅ ✅ ✅ TOTAL COUNT - COUNT ATTENDANCE RECORDS NOT EMPLOYEES
    const countQuery = `
        SELECT COUNT(*) as totalCount
        FROM attendancetable at WITH (NOLOCK)
        INNER JOIN employeemaster em WITH (NOLOCK)
          ON LTRIM(RTRIM(at.emp_code)) = LTRIM(RTRIM(em.empcode))
        WHERE CONVERT(date, at.dateoffice) BETWEEN '${startDate}' AND '${endDate}'
          AND ${filterConditions}
      `;

    const countResult = await sequelize.query(countQuery);
    const totalRecords = countResult[0][0]?.totalCount || 0;

    if (totalRecords === 0) {
      return res.status(200).json({
        success: true,
        data: [],
        totalRecords: 0,
        totalPages: 0,
        currentPage: page,
        pageSize: limit,
        message: "✅ No attendance records found",
      });
    }

    // ✅ ✅ ✅ PAGINATION - WITH OFFSET/LIMIT ON ATTENDANCE RECORDS
    const offset = (page - 1) * limit;

    // ✅ Step-2 main data WITH PAGINATION
    const mainQuery = `
        SELECT
          LTRIM(RTRIM(at.emp_code)) AS Empcode,
          LTRIM(RTRIM(em.empcode)) AS EmployeeCode,
          CONCAT(
            ISNULL(em.title, ''),
            CASE WHEN ISNULL(em.title, '') != '' THEN ' ' ELSE '' END,
            ISNULL(em.empfirstname, ''), ' ', ISNULL(em.EMPLASTNAME, '')
          ) AS EmployeeName,
          em.CURRENTJOINDATE AS JoiningDate,
          ISNULL(m1.misc_name, 'N/A') AS Region,
          ISNULL(m2.misc_name, 'N/A') AS Channel,
          ISNULL(m3.misc_name, 'N/A') AS Cluster,
          ISNULL(m4.misc_name, 'N/A') AS Location,
          ISNULL(m5.misc_name, 'N/A') AS Section,
          ISNULL(m6.misc_name, 'N/A') AS Department,
          em.EMPLOYEEDESIGNATION AS Designation,
          CONVERT(varchar(10), CONVERT(date, at.dateoffice), 23) AS AttendanceDate,
          LTRIM(RTRIM(ISNULL(at.status, 'A'))) AS Status,
          CONVERT(VARCHAR(8), ISNULL(at.in1, '00:00:00'), 108) AS InTime,
          CONVERT(VARCHAR(8), ISNULL(at.out1, '00:00:00'), 108) AS OutTime,
          ISNULL(at.In_Photo, '') AS InPhoto,
          ISNULL(at.Out_Photo, '') AS OutPhoto
        FROM attendancetable at WITH (NOLOCK)
        INNER JOIN employeemaster em WITH (NOLOCK)
          ON LTRIM(RTRIM(at.emp_code)) = LTRIM(RTRIM(em.empcode))
        LEFT JOIN misc_mst m1 WITH (NOLOCK) ON m1.misc_type = 91  AND m1.misc_code = em.sal_region
        LEFT JOIN misc_mst m2 WITH (NOLOCK) ON m2.misc_type = 627 AND m2.misc_code = em.channel
        LEFT JOIN misc_mst m3 WITH (NOLOCK) ON m3.misc_type = 626 AND m3.misc_code = em.cluster
        LEFT JOIN misc_mst m4 WITH (NOLOCK) ON m4.misc_type = 85  AND m4.misc_code = em.location
        LEFT JOIN misc_mst m5 WITH (NOLOCK) ON m5.misc_type = 81  AND m5.misc_code = em.section
        LEFT JOIN misc_mst m6 WITH (NOLOCK) ON m6.misc_type = 68  AND m6.misc_code = em.division
        WHERE CONVERT(date, at.dateoffice) BETWEEN '${startDate}' AND '${endDate}'
          AND ${filterConditions}
        ORDER BY LTRIM(RTRIM(at.emp_code)) ASC, CONVERT(date, at.dateoffice) ASC
        OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY
      `;

    const result = await sequelize.query(mainQuery);
    let allResults = result[0] || [];

    if (!allResults.length) {
      return res.status(200).json({
        success: true,
        data: [],
        totalRecords: totalRecords,
        totalPages: Math.ceil(totalRecords / limit),
        currentPage: page,
        pageSize: limit,
        recordsOnThisPage: 0,
        message: "✅ No attendance records found on this page",
      });
    }

    // ✅ Step-3 images (same keys)
    const imageQuery = `
        SELECT
          LTRIM(RTRIM(emp_code)) AS emp_code,
          CONVERT(varchar(10), CONVERT(date, dateoffice), 23) AS AttendanceDate,
          In_Photo,
          Out_Photo
        FROM attendancetable WITH (NOLOCK)
        WHERE CONVERT(date, dateoffice) BETWEEN '${startDate}' AND '${endDate}'
          AND (
            (In_Photo IS NOT NULL AND LEN(LTRIM(RTRIM(In_Photo))) > 0)
            OR
            (Out_Photo IS NOT NULL AND LEN(LTRIM(RTRIM(Out_Photo))) > 0)
          )
      `;

    const images = await sequelize.query(imageQuery);
    const rows = images[0] || [];

    if (rows.length) {
      const imageMap = {};
      rows.forEach((img) => {
        imageMap[`${img.emp_code}_${img.AttendanceDate}`] = img;
      });

      allResults = allResults.map((row) => {
        const key = `${row.Empcode}_${row.AttendanceDate}`;
        const hit = imageMap[key];
        if (hit) {
          if (hit.In_Photo?.trim()) {
            row.InPhotoUrl =
              "https://erp.autovyn.com/backend/fetch?filePath=" + hit.In_Photo;
          }
          if (hit.Out_Photo?.trim()) {
            row.OutPhotoUrl =
              "https://erp.autovyn.com/backend/fetch?filePath=" + hit.Out_Photo;
          }
        }
        return row;
      });
    }

    const totalPages = Math.ceil(totalRecords / limit);

    return res.status(200).json({
      success: true,
      data: allResults,
      pagination: {
        totalRecords: totalRecords,
        totalPages,
        currentPage: page,
        pageSize: limit,
        hasNextPage: page < totalPages,
        hasPreviousPage: page > 1,
        recordsOnThisPage: allResults.length,
      },
      dateRange: { from: startDate, to: endDate },
      message: `✅ Successfully retrieved ${allResults.length} attendance records (Page ${page}/${totalPages})`,
    });
  } catch (error) {
    return res.status(500).json({
      success: false,
      message: error.message,
      data: [],
      pagination: {
        totalRecords: 0,
        totalPages: 0,
        currentPage: 1,
        pageSize: 10,
        hasNextPage: false,
        hasPreviousPage: false,
      },
    });
  } finally {
    if (sequelize) await sequelize.close();
  }
};

exports.getInsuranceExecutives = async function (req, res) {
  let sequelize;

  try {
    sequelize = await dbname(req, req.headers.compcode);

    const { search, Loc_Code } = req.body || {};

    // ── Base WHERE — sirf active employees ──
    let whereConditions = `WHERE LASTWOR_DATE IS NULL`;
    const replacements = {};

    // ── Loc_Code filter (optional) ────────────────────────────────
    if (Loc_Code) {
      const locArr = Array.isArray(Loc_Code)
        ? Loc_Code.flat()
          .map((v) => String(v).trim())
          .filter(Boolean)
        : String(Loc_Code)
          .trim()
          .split(/[,\s]+/)
          .map((x) => x.trim())
          .filter(Boolean);

      if (locArr.length === 1) {
        whereConditions += ` AND Loc_Code = :Loc_Code`;
        replacements.Loc_Code = locArr[0];
      } else if (locArr.length > 1) {
        whereConditions += ` AND Loc_Code IN (:Loc_Codes)`;
        replacements.Loc_Codes = locArr;
      }
    }

    // ── Search filter ─────────────────────────────────────────────
    // if (search) {
    //   const raw = String(search).trim();
    //   const searchLike = `%${raw}%`;

    //   whereConditions += `
    //     AND (
    //       UPPER(EMPCODE)                         LIKE UPPER(:search)
    //       OR UPPER(ISNULL(EMPFIRSTNAME, ''))     LIKE UPPER(:search)
    //       OR UPPER(ISNULL(EMPLASTNAME,  ''))     LIKE UPPER(:search)
    //       OR UPPER(ISNULL(MOBILENO,     ''))     LIKE UPPER(:search)
    //       OR UPPER(
    //            ISNULL(EMPFIRSTNAME, '') + ' ' +
    //            ISNULL(EMPLASTNAME,  '')
    //          )                                   LIKE UPPER(:search)
    //     )`;
    //   replacements.search = searchLike;
    // }

    // ── DATA query ────────────────────────────────────────────────
    const data = await sequelize.query(
      `SELECT
         EMPCODE,
         ISNULL(EMPFIRSTNAME, '')  AS EMPFIRSTNAME,
         ISNULL(EMPLASTNAME,  '')  AS EMPLASTNAME,
         ISNULL(
           LTRIM(RTRIM(ISNULL(EMPFIRSTNAME, ''))) + ' ' +
           LTRIM(RTRIM(ISNULL(EMPLASTNAME,  ''))),
           ''
         )                         AS FULL_NAME,
         ISNULL(MOBILENO, '')      AS MOBILENO
       FROM dbo.EMPLOYEEMASTER
       ${whereConditions}
       ORDER BY EMPCODE ASC`,
      {
        replacements,
        type: QueryTypes.SELECT,
      },
    );

    return res.status(200).send({
      success: true,
      totalRecords: data.length,
      data,
    });
  } catch (error) {
    console.error("Get Insurance Executives Error:", error);
    return res.status(500).send({
      success: false,
      message: "Internal Server Error",
      error: error.message,
    });
  } finally {
    if (sequelize) {
      await sequelize.close();
    }
  }
};

exports.transferInsuranceWorkload = async function (req, res) {
  let sequelize;
  let transaction;

  try {
    sequelize = await dbname(req, req.headers.compcode);

    const body = req.body || {};

    // ════════════════════════════════════════════════════════
    // VALIDATION
    // ════════════════════════════════════════════════════════
    if (!body.from_Emp_Code && !body.from_exec_mobile) {
      return res.status(400).send({
        success: false,
        message: "from_Emp_Code or from_exec_mobile is required",
      });
    }

    if (!body.to_Emp_Code) {
      return res.status(400).send({
        success: false,
        message: "to_Emp_Code is required",
      });
    }

    if (
      body.from_Emp_Code &&
      body.to_Emp_Code &&
      String(body.from_Emp_Code).trim() === String(body.to_Emp_Code).trim()
    ) {
      return res.status(400).send({
        success: false,
        message: "from_Emp_Code and to_Emp_Code cannot be the same",
      });
    }

    if (!body.Loc_Code) {
      return res.status(400).send({
        success: false,
        message: "Loc_Code is required",
      });
    }

    // ✅ Selected UTDs validate karo
    const selectedInsuRenewalUTDs = Array.isArray(body.selectedInsuRenewalUTDs)
      ? body.selectedInsuRenewalUTDs
        .map((v) => Number(v))
        .filter((v) => !isNaN(v) && v > 0)
      : [];

    if (selectedInsuRenewalUTDs.length === 0) {
      return res.status(400).send({
        success: false,
        message:
          "selectedInsuRenewalUTDs is required — please select at least one insurance renewal",
      });
    }

    // ── Normalize ──────────────────────────────────────────
    const fromEmpCode = body.from_Emp_Code
      ? String(body.from_Emp_Code).trim()
      : null;

    const toEmpCode = String(body.to_Emp_Code).trim();
    const toExecName = body.to_exec_name
      ? String(body.to_exec_name).trim()
      : null;
    const toExecMobile = body.to_exec_mobile
      ? String(body.to_exec_mobile).trim()
      : null;

    if (toExecMobile && !/^[0-9]{10,15}$/.test(toExecMobile)) {
      return res.status(400).send({
        success: false,
        message: "to_exec_mobile must contain 10 to 15 digits",
      });
    }

    const locArr = Array.isArray(body.Loc_Code)
      ? body.Loc_Code.flat()
        .map((v) => String(v).trim())
        .filter(Boolean)
      : String(body.Loc_Code)
        .trim()
        .split(/[,\s]+/)
        .map((x) => x.trim())
        .filter(Boolean);

    if (locArr.length === 0) {
      return res.status(400).send({
        success: false,
        message: "Loc_Code is invalid or empty",
      });
    }

    const updatedBy =
      body.Updated_By || req?.user?.UTD || req?.user?.userId || null;

    // ════════════════════════════════════════════════════════
    // STEP 1 — To Executive info fetch karo
    // ════════════════════════════════════════════════════════
    console.log("[INSURANCE TRANSFER] STEP 1: Fetching To Executive...");

    const empResult = await sequelize.query(
      `SELECT TOP 1
         EMPCODE,
         ISNULL(
           LTRIM(RTRIM(ISNULL(EMPFIRSTNAME, ''))) + ' ' +
           LTRIM(RTRIM(ISNULL(EMPLASTNAME,  ''))),
           ''
         ) AS FULL_NAME,
         ISNULL(MOBILENO, '') AS MOBILENO
       FROM dbo.EMPLOYEEMASTER
       WHERE LASTWOR_DATE IS NULL
         AND EMPCODE = :toEmpCode`,
      {
        replacements: { toEmpCode },
        type: QueryTypes.SELECT,
      },
    );

    if (!empResult || empResult.length === 0) {
      return res.status(404).send({
        success: false,
        message: `To Executive with EMPCODE ${toEmpCode} not found or inactive`,
      });
    }

    const toEmployee = empResult[0];

    // ✅ TO EXECUTIVE DATA - यहाँ update होगा
    const finalToEmpCode = String(toEmployee.EMPCODE).trim();
    const finalToExecName =
      toExecName || String(toEmployee.FULL_NAME || "").trim() || "";
    const finalToExecMobile =
      toExecMobile || String(toEmployee.MOBILENO || "").trim() || "";

    console.log("[INSURANCE TRANSFER] ✅ To Executive Found:", {
      finalToEmpCode,
      finalToExecName,
      finalToExecMobile,
    });
    console.log("[INSURANCE TRANSFER] Selected UTDs:", selectedInsuRenewalUTDs);

    // ════════════════════════════════════════════════════════
    // STEP 2 — Verify karo ki selected UTDs executive ke hain
    // ════════════════════════════════════════════════════════
    console.log("[INSURANCE TRANSFER] STEP 2: Verifying records...");

    const verifyReplace = { selectedInsuRenewalUTDs };
    let verifyWhere = `
      WHERE ir.EXPORT_TYPE = 1
        AND ir.UTD IN (:selectedInsuRenewalUTDs)
    `;

    if (fromEmpCode) {
      verifyWhere += ` AND ir.DSC_EMPCODE = :fromEmpCode`;
      verifyReplace.fromEmpCode = fromEmpCode;
    }

    // ✅ LOC_CODE filter - INSU_RENEWAL_MST से (TRAN_ID link करके)
    if (locArr.length === 1) {
      const locValue = parseInt(locArr[0], 10);
      if (!isNaN(locValue)) {
        verifyWhere += ` AND m.LOC_CODE = :Loc_Code`;
        verifyReplace.Loc_Code = locValue;
      }
    } else if (locArr.length > 1) {
      const locInts = locArr
        .map((v) => parseInt(v, 10))
        .filter((v) => !isNaN(v));

      if (locInts.length > 0) {
        verifyWhere += ` AND m.LOC_CODE IN (:Loc_Codes)`;
        verifyReplace.Loc_Codes = locInts;
      }
    }

    console.log("[INSURANCE TRANSFER] Verify SQL:", {
      verifyWhere,
      verifyReplace,
    });

    const verifyResult = await sequelize.query(
      `SELECT COUNT(*) AS total
       FROM dbo.INSU_RENEWAL ir
       LEFT JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = ir.TRAN_ID
       ${verifyWhere}`,
      { replacements: verifyReplace, type: QueryTypes.SELECT },
    );

    const verifiedCount = Number(verifyResult?.[0]?.total || 0);

    console.log("[INSURANCE TRANSFER] ✅ Verified Count:", verifiedCount);

    if (verifiedCount === 0) {
      return res.status(404).send({
        success: false,
        message:
          "No matching insurance renewal records found for selected tasks and executive",
        debug: { selectedInsuRenewalUTDs, fromEmpCode, Loc_Code: locArr },
      });
    }

    // ════════════════════════════════════════════════════════
    // STEP 3 — Transaction: Update selected records only
    // ════════════════════════════════════════════════════════
    console.log("[INSURANCE TRANSFER] STEP 3: Starting transaction...");

    transaction = await sequelize.transaction();

    // ✅ UPDATE: INSU_RENEWAL — sirf selected UTDs
    const updateInsuReplace = {
      DSC_EMPCODE: finalToEmpCode, // ✅ Employee Code
      DSC_NAME: finalToExecName, // ✅ Employee Name
      DSC_MOB_NO: finalToExecMobile, // ✅ Employee Mobile
      Updated_By: updatedBy,
      selectedInsuRenewalUTDs,
    };

    let updateInsuWhere = `WHERE EXPORT_TYPE = 1 AND UTD IN (:selectedInsuRenewalUTDs)`;

    if (fromEmpCode) {
      updateInsuWhere += ` AND DSC_EMPCODE = :fromEmpCode`;
      updateInsuReplace.fromEmpCode = fromEmpCode;
    }

    console.log("[INSURANCE TRANSFER] Update SQL:", {
      updateInsuWhere,
      updateInsuReplace,
    });

    // ✅ Execute UPDATE query
    await sequelize.query(
      `UPDATE dbo.INSU_RENEWAL
       SET
         DSC_EMPCODE   = :DSC_EMPCODE,
         DSC_NAME      = :DSC_NAME,
         DSC_MOB_NO    = :DSC_MOB_NO,
         CREATED_AT    = GETDATE()
       ${updateInsuWhere}`,
      {
        replacements: updateInsuReplace,
        type: QueryTypes.UPDATE,
        transaction,
      },
    );

    console.log(
      `[INSURANCE TRANSFER] ✅ INSU_RENEWAL updated | ${selectedInsuRenewalUTDs.length} records`,
    );

    // ✅ Commit transaction
    await transaction.commit();

    console.log("[INSURANCE TRANSFER] ✅ Transaction committed successfully!");

    // ════════════════════════════════════════════════════════
    // RESPONSE
    // ════════════════════════════════════════════════════════
    return res.status(200).send({
      success: true,
      message: `${selectedInsuRenewalUTDs.length} insurance renewal(s) transferred successfully`,
      transferredRecords: selectedInsuRenewalUTDs.length,
      from: {
        emp_code: fromEmpCode || null,
      },
      to: {
        emp_code: finalToEmpCode,
        exec_name: finalToExecName,
        exec_mobile: finalToExecMobile,
      },
      transferredInsuRenewalUTDs: selectedInsuRenewalUTDs,
      Loc_Code: locArr,
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    console.error("[INSURANCE TRANSFER] ❌ Error occurred:", error.message);

    if (transaction) {
      try {
        await transaction.rollback();
        console.log(
          "[INSURANCE TRANSFER] Transaction rolled back successfully",
        );
      } catch (rollbackError) {
        console.error(
          "[INSURANCE TRANSFER] Rollback error:",
          rollbackError.message,
        );
      }
    }

    console.error("Transfer Insurance Workload Error:", error);
    return res.status(500).send({
      success: false,
      message: "Internal Server Error",
      error: error.message,
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
        console.log("[INSURANCE TRANSFER] Database connection closed");
      } catch (closeError) {
        console.error(
          "[INSURANCE TRANSFER] Connection close error:",
          closeError.message,
        );
      }
    }
  }
};

exports.getPendingInsuranceByExecutive = async function (req, res) {
  let sequelize;

  try {
    sequelize = await dbname(req, req.headers.compcode);

    const body = req.body || {};

    if (!body.from_Emp_Code && !body.from_exec_mobile) {
      return res.status(400).send({
        success: false,
        message: "from_Emp_Code or from_exec_mobile is required",
      });
    }

    if (!body.Loc_Code) {
      return res.status(400).send({
        success: false,
        message: "Loc_Code is required",
      });
    }

    const locArr = Array.isArray(body.Loc_Code)
      ? body.Loc_Code.flat()
        .map((v) => String(v).trim())
        .filter(Boolean)
      : String(body.Loc_Code)
        .trim()
        .split(/[,\s]+/)
        .map((x) => x.trim())
        .filter(Boolean);

    const whereConditions = [];
    const replacements = {};

    // ✅ Base conditions
    whereConditions.push(`ir.EXPORT_TYPE = 1`);

    // ✅ From Employee filter
    if (body.from_Emp_Code) {
      whereConditions.push(`ir.DSC_EMPCODE = :fromEmpCode`);
      replacements.fromEmpCode = String(body.from_Emp_Code).trim();
    } else if (body.from_exec_mobile) {
      whereConditions.push(
        `LTRIM(RTRIM(ISNULL(ir.DSC_MOB_NO, ''))) = :fromMobile`,
      );
      replacements.fromMobile = String(body.from_exec_mobile).trim();
    }

    // ✅ NULL check
    whereConditions.push(`ir.POLICY_END_DATE IS NOT NULL`);

    // ✅ LOC_CODE filter - sirf INSU_RENEWAL table se (agar wahan hai)
    // Agar LOC_CODE INSU_RENEWAL mein nahi hai to yeh line remove karo
    // if (locArr.length === 1) {
    //   whereConditions.push(`ir.LOC_CODE = :Loc_Code`);
    //   replacements.Loc_Code = parseInt(locArr[0], 10);
    // }

    const whereClause = whereConditions.join(" AND ");

    console.log("[getPendingInsurance] 🔍 Query:", {
      whereClause,
      replacements,
    });

    const data = await sequelize.query(
      `SELECT
         ir.UTD                                    AS Insurance_UTD,
         ir.CUST_NAME,
         ir.CUST_MOB_NO,
         ir.POLICY_NAME,
         ir.POLICY_NUMBER,
         ir.VEHICAL_REG_NO,
         ir.MODEL_NAME,
         CONVERT(varchar(10), ir.POLICY_START_DATE, 23) AS POLICY_START_DATE,
         CONVERT(varchar(10), ir.POLICY_END_DATE, 23)   AS POLICY_END_DATE,
         ir.DSC_EMPCODE,
         ir.DSC_NAME,
         ir.DSC_MOB_NO,
         DATEDIFF(DAY, GETDATE(), ir.POLICY_END_DATE)    AS DAYS_REMAINING
       FROM dbo.INSU_RENEWAL ir
       WHERE ${whereClause}
       ORDER BY ir.POLICY_END_DATE ASC`,
      { replacements, type: QueryTypes.SELECT },
    );

    console.log("[getPendingInsurance] ✅ Success | Records:", data.length);

    return res.status(200).send({
      success: true,
      totalRecords: data.length,
      data,
    });
  } catch (error) {
    console.error("[getPendingInsurance] ❌ Error:", {
      message: error.message || "Unknown error",
      sql: error.sql,
    });

    return res.status(500).send({
      success: false,
      message: error.message || "Internal Server Error",
      error: error.message,
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};

exports.getEmployeeAnniversaries = async function (req, res) {
  let sequelize;

  try {
    sequelize = await dbname(req, req.headers.compcode);

    const { search, Loc_Code } = req.body || {};

    // ── Base WHERE — sirf active employees ──
    let whereConditions = `WHERE LASTWOR_DATE IS NULL`;
    const replacements = {};

    // ── Loc_Code filter (optional) ────────────────────────────────
    // if (Loc_Code) {
    //   const locArr = Array.isArray(Loc_Code)
    //     ? Loc_Code.flat().map((v) => String(v).trim()).filter(Boolean)
    //     : String(Loc_Code)
    //         .trim()
    //         .split(/[,\s]+/)
    //         .map((x) => x.trim())
    //         .filter(Boolean);

    //   if (locArr.length === 1) {
    //     whereConditions += ` AND Loc_Code = :Loc_Code`;
    //     replacements.Loc_Code = locArr[0];
    //   } else if (locArr.length > 1) {
    //     whereConditions += ` AND Loc_Code IN (:Loc_Codes)`;
    //     replacements.Loc_Codes = locArr;
    //   }
    // }

    // ── Search filter ─────────────────────────────────────────────
    // if (search) {
    //   const raw = String(search).trim();
    //   const searchLike = `%${raw}%`;

    //   whereConditions += `
    //     AND (
    //       UPPER(EMPCODE)                         LIKE UPPER(:search)
    //       OR UPPER(ISNULL(EMPFIRSTNAME, ''))     LIKE UPPER(:search)
    //       OR UPPER(ISNULL(EMPLASTNAME,  ''))     LIKE UPPER(:search)
    //       OR UPPER(ISNULL(MOBILENO,     ''))     LIKE UPPER(:search)
    //       OR UPPER(
    //            ISNULL(EMPFIRSTNAME, '') + ' ' +
    //            ISNULL(EMPLASTNAME,  '')
    //          )                                   LIKE UPPER(:search)
    //     )`;
    //   replacements.search = searchLike;
    // }

    // ── Get today's date ──
    const today = new Date();
    const todayMonth = String(today.getMonth() + 1).padStart(2, "0");
    const todayDate = String(today.getDate()).padStart(2, "0");
    const todayMMDD = `${todayMonth}${todayDate}`;

    // ── DATA query with anniversary details ────────────────────────
    const data = await sequelize.query(
      `SELECT
         EMPCODE,
         ISNULL(EMPFIRSTNAME, '')  AS EMPFIRSTNAME,
         ISNULL(EMPLASTNAME, '')   AS EMPLASTNAME,
         ISNULL(
           LTRIM(RTRIM(ISNULL(EMPFIRSTNAME, ''))) + ' ' +
           LTRIM(RTRIM(ISNULL(EMPLASTNAME, ''))),
           ''
         )                         AS FULL_NAME,
         ISNULL(MOBILENO, '')      AS MOBILENO,
         ISNULL(MOBILE_NO, '')     AS MOBILE_NO_ALT,
         ISNULL(LANDLINENO, '')    AS LANDLINENO,
         ISNULL(EMPLOYEEDESIGNATION, '')   AS DESIGNATION,
         ISNULL(ROLE, '')          AS ROLE,
         ISNULL(LOCATION, '')      AS LOCATION,
         ISNULL(GRADE, '')         AS GRADE,
         ISNULL(DIVISION, '')      AS DIVISION,
         ISNULL(UNIT, '')          AS UNIT,
         ISNULL(SECTION, '')       AS SECTION,
         DOB,
         CURRENTJOINDATE           AS DOJ,
         DOM,
         ISNULL(GENDER, '')        AS GENDER,
         ISNULL(MARITALSTATUS, '') AS MARITALSTATUS,
         ISNULL(BLOODGROUP, '')    AS BLOODGROUP,
         ISNULL(PANNO, '')         AS PANNO,
         ISNULL(ADHARNO, '')       AS ADHARNO,
         ISNULL(CORPORATEMAILID, '')   AS EMAIL,
         ISNULL(BANKNAME, '')      AS BANKNAME,
         ISNULL(BANKACCOUNTNO, '') AS BANKACCOUNTNO,
         ISNULL(IFSC_CODE, '')     AS IFSC_CODE,
         CASE
           WHEN FORMAT(DOB, 'MMdd') = :todayMMDD 
           THEN 1 
           ELSE 0 
         END AS IS_BIRTHDAY,
         CASE
           WHEN FORMAT(CURRENTJOINDATE, 'MMdd') = :todayMMDD 
           THEN 1 
           ELSE 0 
         END AS IS_WORK_ANNIVERSARY,
         CASE
           WHEN FORMAT(DOM, 'MMdd') = :todayMMDD 
           THEN 1 
           ELSE 0 
         END AS IS_MARRIAGE_ANNIVERSARY,
         CASE
           WHEN DOB IS NOT NULL AND FORMAT(DOB, 'MMdd') = :todayMMDD
           THEN DATEDIFF(YEAR, DOB, GETDATE())
           ELSE NULL
         END AS BIRTHDAY_AGE,
         CASE
           WHEN CURRENTJOINDATE IS NOT NULL AND FORMAT(CURRENTJOINDATE, 'MMdd') = :todayMMDD
           THEN DATEDIFF(YEAR, CURRENTJOINDATE, GETDATE())
           ELSE NULL
         END AS WORK_ANNIVERSARY_YEARS,
         CASE
           WHEN DOM IS NOT NULL AND FORMAT(DOM, 'MMdd') = :todayMMDD
           THEN DATEDIFF(YEAR, DOM, GETDATE())
           ELSE NULL
         END AS MARRIAGE_ANNIVERSARY_YEARS
       FROM dbo.EMPLOYEEMASTER
       ${whereConditions}
       ORDER BY 
         IS_BIRTHDAY DESC,
         IS_WORK_ANNIVERSARY DESC,
         IS_MARRIAGE_ANNIVERSARY DESC,
         EMPCODE ASC`,
      {
        replacements: {
          ...replacements,
          todayMMDD,
        },
        type: QueryTypes.SELECT,
      },
    );

    // ── Count anniversaries ────────────────────────────────────────
    const counts = {
      totalRecords: data.length,
      birthdays: data.filter((emp) => emp.IS_BIRTHDAY === 1).length,
      workAnniversaries: data.filter((emp) => emp.IS_WORK_ANNIVERSARY === 1)
        .length,
      marriageAnniversaries: data.filter(
        (emp) => emp.IS_MARRIAGE_ANNIVERSARY === 1,
      ).length,
      todayCount:
        data.filter(
          (emp) =>
            emp.IS_BIRTHDAY === 1 ||
            emp.IS_WORK_ANNIVERSARY === 1 ||
            emp.IS_MARRIAGE_ANNIVERSARY === 1,
        ).length || 0,
    };

    // ── Separate data by anniversary type ──
    const separatedData = {
      birthdays: data
        .filter((emp) => emp.IS_BIRTHDAY === 1)
        .map((emp) => ({
          empCode: emp.EMPCODE,
          firstName: emp.EMPFIRSTNAME,
          lastName: emp.EMPLASTNAME,
          mobileNo: emp.MOBILENO || emp.MOBILE_NO_ALT,
          designation: emp.DESIGNATION,
          role: emp.ROLE,
          location: emp.LOCATION,
          dob: emp.DOB,
          gender: emp.GENDER,
          age: emp.BIRTHDAY_AGE,
        })),
      workAnniversaries: data
        .filter((emp) => emp.IS_WORK_ANNIVERSARY === 1)
        .map((emp) => ({
          empCode: emp.EMPCODE,
          firstName: emp.EMPFIRSTNAME,
          lastName: emp.EMPLASTNAME,
          mobileNo: emp.MOBILENO || emp.MOBILE_NO_ALT,
          designation: emp.DESIGNATION,
          role: emp.ROLE,
          location: emp.LOCATION,

          doj: emp.DOJ,

          yearsOfService: emp.WORK_ANNIVERSARY_YEARS,
        })),
      marriageAnniversaries: data
        .filter((emp) => emp.IS_MARRIAGE_ANNIVERSARY === 1)
        .map((emp) => ({
          empCode: emp.EMPCODE,
          firstName: emp.EMPFIRSTNAME,
          lastName: emp.EMPLASTNAME,
          mobileNo: emp.MOBILENO || emp.MOBILE_NO_ALT,
          designation: emp.DESIGNATION,
          role: emp.ROLE,
          location: emp.LOCATION,
          grade: emp.GRADE,
          maritalStatus: emp.MARITALSTATUS,
          dom: emp.DOM,
          anniversaryType: "Marriage Anniversary",
          yearsOfMarriage: emp.MARRIAGE_ANNIVERSARY_YEARS,
        })),
    };

    return res.status(200).send({
      success: true,
      timestamp: today,
      counts,
      separatedData,
      // allData: data,
    });
  } catch (error) {
    console.error("Get Employee Anniversaries Error:", error);
    return res.status(500).send({
      success: false,
      message: "Internal Server Error",
      error: error.message,
    });
  } finally {
    if (sequelize) {
      await sequelize.close();
    }
  }
};

exports.assignInsuranceRenewalsToDSE = async function (req, res) {
  let sequelize;

  try {
    sequelize = await dbname(req, req.headers.compcode);

    const { DSE_EMPCODE, selectedInsuRenewalUTDs } = req.body || {};

    // ── Validate inputs ──────────────────────────────────────────
    if (!DSE_EMPCODE || String(DSE_EMPCODE).trim() === "") {
      return res.status(400).send({
        success: false,
        message: "DSE_EMPCODE is required",
      });
    }

    // ✅ Validate selected UTDs
    if (!selectedInsuRenewalUTDs || !Array.isArray(selectedInsuRenewalUTDs) || selectedInsuRenewalUTDs.length === 0) {
      return res.status(400).send({
        success: false,
        message: "Please select at least one insurance renewal",
      });
    }

    const dseCode = String(DSE_EMPCODE).trim();

    console.log("DSE Code:", dseCode);
    console.log("Selected UTDs:", selectedInsuRenewalUTDs);

    // ── Get DSE details from EMPLOYEEMASTER ───────────────────────
    const dseDetails = await sequelize.query(
      `SELECT
          EMPCODE,
          ISNULL(EMPFIRSTNAME, '')  AS EMPFIRSTNAME,
          ISNULL(EMPLASTNAME,  '')  AS EMPLASTNAME,
          ISNULL(
            LTRIM(RTRIM(ISNULL(EMPFIRSTNAME, ''))) + ' ' +
            LTRIM(RTRIM(ISNULL(EMPLASTNAME,  ''))),
            ''
          )                         AS FULL_NAME,
          ISNULL(MOBILENO, '')      AS MOBILENO
        FROM dbo.EMPLOYEEMASTER
        WHERE EMPCODE = :dseCode
        AND LASTWOR_DATE IS NULL`,
      {
        replacements: { dseCode },
        type: QueryTypes.SELECT,
      },
    );

    if (!dseDetails || dseDetails.length === 0) {
      return res.status(404).send({
        success: false,
        message: `DSE with code ${dseCode} not found or inactive`,
      });
    }

    const dse = dseDetails[0];
    console.log("DSE Found:", dse);

    // ✅ STEP 1: UPDATE only SELECTED INSU_RENEWAL records
    // ────────────────────────────────────────────────────────────
    const updateResult = await sequelize.query(
      `UPDATE dbo.INSU_RENEWAL
        SET 
          DSC_EMPCODE = :dseEmpCode,
          DSC_NAME = :dseName,
          DSC_MOB_NO = :dseMobNo
        WHERE UTD IN (:selectedUTDs)
        AND EXPORT_TYPE = 1`,
      {
        replacements: {
          dseEmpCode: dse.EMPCODE,
          dseName: dse.FULL_NAME,
          dseMobNo: dse.MOBILENO,
          selectedUTDs: selectedInsuRenewalUTDs,
        },
        type: QueryTypes.UPDATE,
      },
    );

    const updatedCount = updateResult[1];
    console.log(`Updated ${updatedCount} records`);

    // ✅ STEP 2: Get ONLY the selected and updated records
    // ────────────────────────────────────────────────────────────
    const assignedRenewals = await sequelize.query(
      `SELECT
          ir.UTD,
          ir.CUST_NAME,
          ir.CUST_MOB_NO,
          ir.POLICY_NAME,
          ir.POLICY_NUMBER,
          ir.VEHICAL_REG_NO,
          ir.MODEL_NAME,
          ir.POLICY_START_DATE,
          ir.POLICY_END_DATE,
          ir.EXPORT_TYPE,
          ir.DSC_EMPCODE,
          ir.DSC_NAME,
          ir.DSC_MOB_NO,
          irm.UTD AS MST_UTD,
          irm.LOC_CODE
        FROM dbo.INSU_RENEWAL ir
        INNER JOIN dbo.INSU_RENEWAL_MST irm 
          ON ir.TRAN_ID = irm.UTD
        WHERE ir.UTD IN (:selectedUTDs)
        AND ir.EXPORT_TYPE = 1
        AND ir.DSC_EMPCODE = :dseCode
        ORDER BY ir.VEHICAL_REG_NO, ir.POLICY_END_DATE ASC`,
      {
        replacements: {
          selectedUTDs: selectedInsuRenewalUTDs,
          dseCode: dse.EMPCODE,
        },
        type: QueryTypes.SELECT,
      },
    );

    console.log("Assigned renewals found:", assignedRenewals.length);
    console.log("Sample renewal:", assignedRenewals[0]);

    if (!assignedRenewals || assignedRenewals.length === 0) {
      return res.status(200).send({
        success: true,
        message: "No insurance renewals updated",
        dseDetails: dse,
        assignments: [],
        totalAssignments: 0,
        updatedRecords: updatedCount,
      });
    }

    // ── Group assignments by vehicle ───────────────────────────────
    const assignmentsByVehicle = {};
    assignedRenewals.forEach((renewal) => {
      const vehicleReg = renewal.VEHICAL_REG_NO;
      if (!assignmentsByVehicle[vehicleReg]) {
        assignmentsByVehicle[vehicleReg] = [];
      }
      assignmentsByVehicle[vehicleReg].push(renewal);
    });

    console.log("Grouped vehicles:", Object.keys(assignmentsByVehicle));

    // ── Prepare response ──────────────────────────────────────────
    const finalAssignments = Object.keys(assignmentsByVehicle).map((vehicle) => ({
      VEHICAL_REG_NO: vehicle,
      DSE_EMPCODE: dse.EMPCODE,
      DSE_NAME: dse.FULL_NAME,
      DSE_MOB_NO: dse.MOBILENO,
      ASSIGNED_ON: new Date(),
      CUSTOMERS: assignmentsByVehicle[vehicle].map((cust) => ({
        UTD: cust.UTD,
        CUST_NAME: cust.CUST_NAME,
        CUST_MOB_NO: cust.CUST_MOB_NO,
        POLICY_NAME: cust.POLICY_NAME,
        POLICY_NUMBER: cust.POLICY_NUMBER,
        MODEL_NAME: cust.MODEL_NAME,
        POLICY_START_DATE: cust.POLICY_START_DATE,
        POLICY_END_DATE: cust.POLICY_END_DATE,
        EXPORT_TYPE: cust.EXPORT_TYPE,
        MST_UTD: cust.MST_UTD,
        LOC_CODE: cust.LOC_CODE,
      })),
      CUSTOMER_COUNT: assignmentsByVehicle[vehicle].length,
    }));

    return res.status(200).send({
      success: true,
      message: `${updatedCount} insurance renewals assigned successfully to ${dse.FULL_NAME}`,
      dseDetails: {
        EMPCODE: dse.EMPCODE,
        FULL_NAME: dse.FULL_NAME,
        MOBILENO: dse.MOBILENO,
      },
      updatedRecords: updatedCount,
      totalVehicles: finalAssignments.length,
      totalCustomers: assignedRenewals.length,
      assignments: finalAssignments,
    });
  } catch (error) {
    console.error("Assign Insurance Renewals Error:", error);
    return res.status(500).send({
      success: false,
      message: "Internal Server Error",
      error: error.message,
    });
  } finally {
    if (sequelize) {
      await sequelize.close();
    }
  }
};

exports.getExpiringInsuranceRenewals = async function (req, res) {
  let sequelize;

  try {
    sequelize = await dbname(req, req.headers.compcode);

    const { days = 30, Loc_Code } = req.body || {};

    // ── Validate ──────────────────────────────────────────────────
    if (!Loc_Code) {
      return res.status(400).send({
        success: false,
        message: "Loc_Code is required",
      });
    }

    const locCode = String(Loc_Code).trim();
    const dayThreshold = Number(days) || 30;

    console.log("[getExpiringInsuranceRenewals] Days:",);

    // ── Get ALL expiring customers (not filtered by DSE) ──────────
    const expiringRenewals = await sequelize.query(
      `SELECT
         ir.UTD                                    AS Insurance_UTD,
         ir.CUST_NAME,
         ir.CUST_MOB_NO,
         ir.POLICY_NAME,
         ir.POLICY_NUMBER,
         ir.VEHICAL_REG_NO,
         ir.MODEL_NAME,
         CONVERT(varchar(10), ir.POLICY_START_DATE, 23) AS POLICY_START_DATE,
         CONVERT(varchar(10), ir.POLICY_END_DATE, 23)   AS POLICY_END_DATE,
         ir.DSC_EMPCODE,
         ir.DSC_NAME,
         ir.DSC_MOB_NO,
         DATEDIFF(DAY, GETDATE(), ir.POLICY_END_DATE)    AS DAYS_REMAINING
       FROM dbo.INSU_RENEWAL ir
       INNER JOIN dbo.INSU_RENEWAL_MST irm 
         ON ir.TRAN_ID = irm.UTD
       WHERE ir.EXPORT_TYPE = 1
       AND ir.POLICY_END_DATE IS NOT NULL
       AND DATEDIFF(DAY, GETDATE(), ir.POLICY_END_DATE) <= :dayThreshold
       AND DATEDIFF(DAY, GETDATE(), ir.POLICY_END_DATE) > 0
       ORDER BY ir.POLICY_END_DATE ASC`,
      {
        replacements: {

        },
        type: QueryTypes.SELECT,
      },
    );

    console.log(
      "[getExpiringInsuranceRenewals] ✅ Records found:",
      expiringRenewals.length
    );

    return res.status(200).send({
      success: true,
      message: `${expiringRenewals.length} insurance renewals expiring within days`,
      totalRecords: expiringRenewals.length,
      data: expiringRenewals,
    });
  } catch (error) {
    console.error("[getExpiringInsuranceRenewals] ❌ Error:", error);
    return res.status(500).send({
      success: false,
      message: "Internal Server Error",
      error: error.message,
    });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};


exports.SendInsuranceRenewalWhatsAppToCustomer = async function (req, res) {
  let sequelize;
  try {
    const { tranIds } = req.body; // tranIds = INSU_RENEWAL_MST ki UTDs (Array)

    // 1. Validation
    if (!Array.isArray(tranIds) || tranIds.length === 0) {
      return res.status(400).send({
        success: false,
        message: "tranIds array is required (Pass array of INSU_RENEWAL_MST UTDs)",
      });
    }

    // 2. Database Connection
    sequelize = await dbname(req, req.headers.compcode);

    // 3. Fetch Active Questions from INSU_QUESTIONS Table
    let questionsList = [];
    try {
      questionsList = await sequelize.query(`
        SELECT 
          UTD,
          Question_Text,
          Question_Type,
          PAYMENT_AMOUNT,
          Options,
          Sort_Order,
          Is_Active
        FROM dbo.INSU_QUESTIONS
        WHERE (Is_Active = '1' OR Is_Active = 'true' OR Is_Active = 'Y' OR Is_Active IS NULL)
        ORDER BY Sort_Order ASC, UTD ASC
      `, { type: sequelize.QueryTypes.SELECT });
    } catch (qErr) {
      console.warn("Could not fetch INSU_QUESTIONS (or table empty):", qErr.message);
      questionsList = [];
    }

    const results = [];

    // 4. Loop over each Master UTD (tranId)
    for (const mstUtd of tranIds) {
      try {
        // Query to join INSU_RENEWAL_MST and INSU_RENEWAL (where r.EXPORT_TYPE = 1)
        const [[row]] = await sequelize.query(`
          SELECT TOP 1
            m.UTD AS MST_UTD,
            m.VEHICAL_REG_NO,
            m.LOC_CODE,
            r.UTD AS DETAIL_UTD,
            r.TRAN_ID,
            r.CUST_NAME,
            r.CUST_MOB_NO,
            r.POLICY_NAME,
            r.POLICY_NUMBER,
            r.MODEL_NAME,
            CONVERT(VARCHAR(10), r.POLICY_START_DATE, 105) AS POLICY_START_DATE, -- DD-MM-YYYY
            CONVERT(VARCHAR(10), r.POLICY_END_DATE, 105) AS POLICY_END_DATE,     -- DD-MM-YYYY
            r.DSC_EMPCODE,
            r.DSC_NAME,
            r.DSC_MOB_NO,
            r.EXPORT_TYPE
          FROM dbo.INSU_RENEWAL_MST m
          INNER JOIN dbo.INSU_RENEWAL r 
            ON r.TRAN_ID = m.UTD
          WHERE m.UTD = '${mstUtd}'
            AND r.EXPORT_TYPE = 1
            AND ISNULL(m.EXPORT_TYPE, 0) < 3
          ORDER BY r.UTD DESC
        `);

        if (!row) {
          results.push({
            tranId: mstUtd,
            success: false,
            message: "Record not found in INSU_RENEWAL with EXPORT_TYPE = 1 for this Master UTD",
          });
          continue;
        }

        if (!row.CUST_MOB_NO) {
          results.push({
            tranId: mstUtd,
            regNo: row.VEHICAL_REG_NO,
            success: false,
            message: "Customer mobile number (CUST_MOB_NO) is missing",
          });
          continue;
        }

        const customerName = row.CUST_NAME || "Valued Customer";
        const vehicleNo = row.VEHICAL_REG_NO || "N/A";
        const policyName = row.POLICY_NAME || "Insurance Provider";

        // 5. Reset previous submitted answers for this customer so the new WhatsApp link opens completely fresh
        if (row.CUST_NAME) {
          const safeCustName = String(row.CUST_NAME).trim().replace(/'/g, "''");
          try {
            await sequelize.query(`
              UPDATE dbo.INSU_ANSWERS
              SET EXPORT_TYPE = 33
              WHERE Created_By = '${safeCustName}'
                AND (EXPORT_TYPE = 1 OR EXPORT_TYPE IS NULL)
            `);
          } catch (resetErr) {
            console.warn("Could not archive old INSU_ANSWERS:", resetErr.message);
          }
        }

        // Encode compcode & UTD for Customer Preference Link
        const encodedCompCode = Buffer.from(req.headers.compcode).toString("base64");
        const encodedUTD = Buffer.from(`${row.MST_UTD}`).toString("base64");
        const customerWebLink = `${process.env.BASE_URL || ''}/Crm/InsuranceRenewalCustomerView?compcode=${encodedCompCode}&utd=${encodedUTD}`;

        // 6. WhatsApp Message Parameters (Template: insurance_template)
        // Template Mapping:
        // {{1}} = Customer Name (from INSU_RENEWAL)
        // {{2}} = Vehicle Reg No (from INSU_RENEWAL_MST)
        // {{3}} = Policy / Company Name (from INSU_RENEWAL)
        // {{4}} (or Link) = Customer Preference Link
        const whatsappParameters = [
          { type: "text", text: customerName },   // {{1}} Customer Name
          { type: "text", text: vehicleNo },      // {{2}} Vehicle No
          { type: "text", text: policyName },     // {{3}} Policy Name
          { type: "text", text: customerWebLink } // {{4}} Preference Link
        ];

        // 7. Send WhatsApp Message
        const whatsappResponse = await SendWhatsAppMessgae(
          req.headers.compcode,
          row.CUST_MOB_NO,
          "insurance_template",
          whatsappParameters
        );

        results.push({
          tranId: mstUtd,
          regNo: vehicleNo,
          customerName: customerName,
          policyName: policyName,
          mobile: row.CUST_MOB_NO,
          success: true,
          message: "Insurance Renewal WhatsApp sent successfully",
          customerWebLink,
          questionsCount: questionsList.length,
          whatsappResponse
        });

      } catch (innerErr) {
        console.error(`Error sending insurance message for tranId ${mstUtd}:`, innerErr);
        results.push({
          tranId: mstUtd,
          success: false,
          message: innerErr.message,
        });
      }
    }

    return res.status(200).send({
      success: true,
      totalRequested: tranIds.length,
      results,
    });

  } catch (e) {
    console.error("SendInsuranceRenewalWhatsAppToCustomer ERROR:", e);
    return res.status(500).send({ success: false, message: e.message });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};


// ============================================================
// CUSTOMER-FACING VIEW: Insurance Renewal Preference & Questions Page
// ============================================================
exports.InsuranceRenewalCustomerView = async function (req, res) {
  let sequelize;
  try {
    const rawCompCode = req.query.compcode ? Buffer.from(req.query.compcode, "base64").toString("utf-8") : "";
    const rawUtd = req.query.utd ? Buffer.from(req.query.utd, "base64").toString("utf-8") : "";

    if (!rawCompCode || !rawUtd) {
      return res.send(`
        <!DOCTYPE html>
        <html>
        <head><title>Invalid Link</title><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
        <body style="font-family:Segoe UI, sans-serif;text-align:center;padding:50px;background:#f8fafc;">
          <h2 style="color:#dc2626;">Invalid or Expired Link</h2>
          <p style="color:#4b5563;">Please check your WhatsApp link and try again.</p>
        </body>
        </html>
      `);
    }

    sequelize = await dbname(req, rawCompCode);

    // 1. Fetch Vehicle & Customer Info (Master + Detail where EXPORT_TYPE = 1)
    const [[row]] = await sequelize.query(`
      SELECT TOP 1
        m.UTD AS MST_UTD,
        m.VEHICAL_REG_NO,
        m.LOC_CODE,
        r.UTD AS DETAIL_UTD,
        r.TRAN_ID,
        r.CUST_NAME,
        r.CUST_MOB_NO,
        r.POLICY_NAME,
        r.POLICY_NUMBER,
        r.MODEL_NAME,
        CONVERT(VARCHAR(10), r.POLICY_START_DATE, 105) AS POLICY_START_DATE,
        CONVERT(VARCHAR(10), r.POLICY_END_DATE, 105) AS POLICY_END_DATE,
        r.DSC_NAME,
        r.DSC_MOB_NO
      FROM dbo.INSU_RENEWAL_MST m
      INNER JOIN dbo.INSU_RENEWAL r ON r.TRAN_ID = m.UTD
      WHERE m.UTD = '${rawUtd}'
        AND r.EXPORT_TYPE = 1
        AND ISNULL(m.EXPORT_TYPE, 0) < 3
      ORDER BY r.UTD DESC
    `);

    if (!row) {
      return res.send(`
        <!DOCTYPE html>
        <html>
        <head><title>Record Not Found</title><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
        <body style="font-family:Segoe UI, sans-serif;text-align:center;padding:50px;background:#f8fafc;">
          <h2 style="color:#dc2626;">Renewal Record Not Found</h2>
          <p style="color:#4b5563;">The insurance renewal record you are looking for does not exist or has been removed.</p>
        </body>
        </html>
      `);
    }

    // 2. Fetch Active Questions from INSU_QUESTIONS
    let questions = [];
    try {
      questions = await sequelize.query(`
        SELECT 
          UTD,
          Question_Text,
          Question_Type,
          PAYMENT_AMOUNT,
          Options,
          Sort_Order,
          Is_Active
        FROM dbo.INSU_QUESTIONS
        WHERE (Is_Active = '1' OR Is_Active = 'true' OR Is_Active = 'Y' OR Is_Active IS NULL)
        ORDER BY Sort_Order ASC, UTD ASC
      `, { type: sequelize.QueryTypes.SELECT });
    } catch (qErr) {
      console.warn("Error fetching INSU_QUESTIONS:", qErr.message);
      questions = [];
    }

    const customerName = row.CUST_NAME || "Valued Customer";
    const vehicleNo = row.VEHICAL_REG_NO || "N/A";
    const policyName = row.POLICY_NAME || "Insurance Provider";
    const policyNo = row.POLICY_NUMBER || "N/A";
    const expiryDate = row.POLICY_END_DATE || "Due Soon";
    const modelName = row.MODEL_NAME || "Car";

    // 3. Check if already answered in INSU_ANSWERS (matched by Customer Name)
    let existingAnswers = [];
    try {
      const safeCustName = String(customerName).replace(/'/g, "''");
      existingAnswers = await sequelize.query(`
        SELECT 
          UTD,
          TRAN_ID,
          Answer_Text,
          EXPORT_TYPE,
          CONVERT(VARCHAR(19), CREATED_AT, 120) AS ANSWERED_AT
        FROM dbo.INSU_ANSWERS
        WHERE Created_By = '${safeCustName}'
          AND ISNULL(EXPORT_TYPE, 0) = 1
      `, { type: sequelize.QueryTypes.SELECT });
    } catch (_) {
      existingAnswers = [];
    }

    const isAlreadySubmitted = existingAnswers && existingAnswers.length > 0;
    const answeredMap = {};
    existingAnswers.forEach(a => {
      if (a.TRAN_ID) {
        answeredMap[a.TRAN_ID] = a.Answer_Text || "";
      }
    });

    const statusBadgeText = isAlreadySubmitted ? "✅ Response Submitted" : "⏳ Awaiting Your Choice";
    const statusBadgeClass = isAlreadySubmitted ? "status-submitted" : "status-pending";

    // 4. Render HTML
    res.send(`
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Insurance Renewal - ${vehicleNo}</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
    background: linear-gradient(135deg, #f0fdf4 0%, #e0f2fe 100%);
    min-height: 100vh;
    padding: 24px 12px;
    display: flex;
    justify-content: center;
    align-items: flex-start;
  }
  .container { max-width: 760px; width: 100%; animation: slideUp 0.5s ease-out; }
  @keyframes slideUp { from { opacity: 0; transform: translateY(18px); } to { opacity: 1; transform: translateY(0); } }

  .card {
    background: #ffffff;
    border-radius: 20px;
    box-shadow: 0 20px 45px rgba(15, 23, 42, 0.08);
    padding: 36px 30px;
    border: 1px solid rgba(226, 232, 240, 0.9);
  }

  .header {
    text-align: center;
    padding-bottom: 22px;
    margin-bottom: 24px;
    border-bottom: 2px solid #e2e8f0;
    position: relative;
  }
  .header .icon {
    width: 64px; height: 64px; margin: 0 auto 12px;
    background: linear-gradient(135deg, #0284c7, #0369a1);
    color: #fff; font-size: 32px; display: flex; align-items: center; justify-content: center;
    border-radius: 16px; box-shadow: 0 8px 18px rgba(2, 132, 199, 0.25);
  }
  .header h1 { font-size: 24px; font-weight: 700; color: #0f172a; margin-bottom: 6px; letter-spacing: -0.5px; }
  .header p { font-size: 14px; color: #64748b; }

  .status-badge {
    display: inline-block; padding: 6px 20px; border-radius: 9999px; font-size: 13px; font-weight: 700; margin-top: 10px;
  }
  .status-pending { background: #fef3c7; color: #92400e; border: 1.5px solid #fde68a; }
  .status-submitted { background: #dcfce7; color: #166534; border: 1.5px solid #bbf7d0; }

  /* Info Grid */
  .grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 28px; }
  .info-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 14px; padding: 12px 16px; }
  .info-box.full { grid-column: span 2; }
  .info-box .lbl { font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }
  .info-box .val { font-size: 15px; font-weight: 600; color: #0f172a; word-break: break-word; }
  .info-box .val.highlight { color: #0284c7; font-weight: 700; }
  .info-box .val.danger { color: #dc2626; font-weight: 700; }

  /* Section Titles */
  .section-heading {
    font-size: 16px; font-weight: 700; color: #0f172a; text-transform: uppercase; letter-spacing: 0.5px;
    margin: 24px 0 14px; display: flex; align-items: center; gap: 8px;
  }

  /* Question Cards */
  .question-card {
    background: #ffffff;
    border: 1.5px solid #e2e8f0;
    border-radius: 16px;
    padding: 18px 20px;
    margin-bottom: 16px;
    transition: all 0.2s ease;
  }
  .question-card:hover { border-color: #cbd5e1; box-shadow: 0 4px 12px rgba(0,0,0,0.03); }
  .question-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 14px; }
  .question-title { font-size: 15px; font-weight: 600; color: #1e293b; line-height: 1.4; }
  .q-num { color: #0284c7; font-weight: 700; margin-right: 4px; }
  .amount-badge {
    background: #ecfdf5; color: #047857; font-size: 12px; font-weight: 700;
    padding: 4px 10px; border-radius: 8px; border: 1px solid #a7f3d0; white-space: nowrap;
  }

  /* Options list */
  /* Options list */
  .options-group { display: flex; flex-direction: column; gap: 8px; }
  .option-label {
    display: flex; align-items: center; gap: 10px; padding: 10px 14px;
    background: #f8fafc; border: 1.5px solid #e2e8f0; border-radius: 10px;
    cursor: pointer; font-size: 14px; font-weight: 500; color: #334155;
    transition: all 0.15s ease;
  }
  .option-label:hover { background: #f1f5f9; border-color: #cbd5e1; }
  .option-label input[type="radio"], .option-label input[type="checkbox"] {
    accent-color: #0284c7; width: 16px; height: 16px; cursor: pointer;
  }
  .option-label.selected {
    background: #f0f9ff; border-color: #0284c7; color: #0369a1; font-weight: 600;
  }
  .text-input {
    width: 100%; padding: 10px 14px; border: 1.5px solid #e2e8f0; border-radius: 10px;
    font-size: 14px; font-family: inherit; transition: border-color 0.2s;
  }
  .text-input:focus { outline: none; border-color: #0284c7; background: #fff; }

  .other-text-input {
    width: 100%; padding: 10px 14px; margin-top: 8px; border: 1.5px solid #0284c7;
    border-radius: 10px; font-size: 14px; font-family: inherit; background: #ffffff;
    box-shadow: 0 2px 6px rgba(2, 132, 199, 0.08); transition: all 0.2s ease;
  }
  .other-text-input:focus {
    outline: none; border-color: #0369a1; box-shadow: 0 0 0 3px rgba(2, 132, 199, 0.15);
  }

  /* Submit Button */
  .submit-btn {
    width: 100%; padding: 16px 24px; border: none; border-radius: 14px;
    background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
    color: #ffffff; font-size: 16px; font-weight: 700; cursor: pointer;
    margin-top: 20px; box-shadow: 0 10px 20px rgba(2, 132, 199, 0.25);
    transition: transform 0.15s, opacity 0.15s;
  }
  .submit-btn:hover { transform: translateY(-1px); opacity: 0.95; }
  .submit-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; box-shadow: none; }

  /* Notification message */
  #statusMsg {
    margin-top: 16px; text-align: center; font-weight: 600; padding: 12px 16px;
    border-radius: 12px; display: none; font-size: 14px;
  }
  #statusMsg.success { display: block; background: #dcfce7; color: #166534; border: 1.5px solid #86efac; }
  #statusMsg.error { display: block; background: #fee2e2; color: #991b1b; border: 1.5px solid #fca5a5; }
  #statusMsg.loading { display: block; background: #fef3c7; color: #92400e; border: 1.5px solid #fcd34d; }

  .footer { text-align: center; margin-top: 24px; font-size: 12px; color: #94a3b8; }

  @media (max-width: 640px) {
    .card { padding: 22px 16px; }
    .grid { grid-template-columns: 1fr; }
    .info-box.full { grid-column: span 1; }
  }
</style>
</head>
<body>

<div class="container">
  <div class="card">
    <div class="header">
      <div class="icon">🛡️</div>
      <h1>Insurance Renewal Preference</h1>
      <p>Please review your vehicle details & select your renewal preference</p>
      <div class="status-badge ${statusBadgeClass}">${statusBadgeText}</div>
    </div>

    <!-- Vehicle & Customer Details -->
    <div class="grid">
      <div class="info-box">
        <div class="lbl">👤 Customer Name</div>
        <div class="val">${customerName}</div>
      </div>
      <div class="info-box">
        <div class="lbl">🚗 Vehicle Reg No</div>
        <div class="val highlight">${vehicleNo}</div>
      </div>
      <div class="info-box">
        <div class="lbl">🚘 Car Model</div>
        <div class="val">${modelName}</div>
      </div>
      <div class="info-box">
        <div class="lbl">🏢 Current Policy / Insurer</div>
        <div class="val">${policyName}</div>
      </div>
      <div class="info-box">
        <div class="lbl">📄 Policy Number</div>
        <div class="val">${policyNo}</div>
      </div>
      <div class="info-box">
        <div class="lbl">📅 Expiry / Due Date</div>
        <div class="val danger">${expiryDate}</div>
      </div>
    </div>

    <!-- Questions Form -->
    <form id="renewalForm">
      <div class="section-heading">📋 Choose Your Preferences</div>

      ${questions.length === 0 ? `
        <div style="text-align:center;padding:24px;color:#64748b;background:#f8fafc;border-radius:12px;">
          No survey questions configured. Please add any special remarks below.
        </div>
      ` : questions.map((q, idx) => {
      const qId = q.UTD;
      const qType = String(q.Question_Type || 'radio').toLowerCase();
      const savedVal = answeredMap[qId] || '';
      const rawOptions = (q.Options || "Yes,No").split(",").map(o => o.trim()).filter(Boolean);
      const amount = Number(q.PAYMENT_AMOUNT || 0);

      const hasOtherOpt = rawOptions.some(o => o.toLowerCase().startsWith('other'));
      const isOtherChecked = hasOtherOpt && (savedVal.toLowerCase().startsWith('other') || (savedVal && !rawOptions.includes(savedVal) && qType !== 'text'));
      const otherSavedVal = isOtherChecked ? (savedVal.replace(/^Other:\s*/i, '').replace(/^Other/i, '').trim()) : '';

      return `
          <div class="question-card" data-qid="${qId}" data-qtype="${qType}" data-qtext="${escapeHtml(q.Question_Text || '')}">
            <div class="question-header">
              <div class="question-title">
                <span class="q-num">Q${idx + 1}.</span> ${q.Question_Text || 'Question'}
              </div>
              ${amount > 0 ? `<div class="amount-badge">₹ ${amount.toLocaleString('en-IN')}</div>` : ''}
            </div>

            <div class="options-group">
              ${qType === 'text' ? `
                <input 
                  type="text" 
                  class="text-input" 
                  name="q_${qId}" 
                  value="${escapeHtml(savedVal)}" 
                  placeholder="Your answer..." 
                  ${isAlreadySubmitted ? 'disabled' : ''} 
                />
              ` : qType === 'checkbox' ? (
          rawOptions.map((opt, optIdx) => {
            const isThisOther = opt.toLowerCase().startsWith('other');
            const isChecked = isThisOther ? isOtherChecked : savedVal.split(',').map(s => s.trim()).includes(opt);
            return `
                    <label class="option-label ${isChecked ? 'selected' : ''}">
                      <input 
                        type="checkbox" 
                        name="q_${qId}" 
                        value="${escapeHtml(opt)}" 
                        ${isChecked ? 'checked' : ''} 
                        ${isAlreadySubmitted ? 'disabled' : ''}
                        onchange="handleCheckboxSelect(this)"
                      />
                      <span>${opt}</span>
                    </label>
                  `;
          }).join('')
        ) : (
          rawOptions.map((opt, optIdx) => {
            const isThisOther = opt.toLowerCase().startsWith('other');
            const isChecked = isThisOther ? isOtherChecked : ((savedVal === opt) || (!savedVal && optIdx === 0 && !isAlreadySubmitted));
            return `
                    <label class="option-label ${isChecked ? 'selected' : ''}">
                      <input 
                        type="radio" 
                        name="q_${qId}" 
                        value="${escapeHtml(opt)}" 
                        ${isChecked ? 'checked' : ''} 
                        ${isAlreadySubmitted ? 'disabled' : ''}
                        onchange="handleRadioSelect(this)"
                      />
                      <span>${opt}</span>
                    </label>
                  `;
          }).join('')
        )}

              ${hasOtherOpt ? `
                <input 
                  type="text" 
                  class="other-text-input" 
                  placeholder="Please specify your details / reason..."
                  value="${escapeHtml(otherSavedVal)}"
                  style="display: ${isOtherChecked ? 'block' : 'none'};"
                  ${isAlreadySubmitted ? 'disabled' : ''}
                />
              ` : ''}
            </div>
          </div>
        `;
    }).join('')}
      
      ${!isAlreadySubmitted ? `
        <button type="button" class="submit-btn" id="submitBtn" onclick="submitPreferences()">
          ✅ Submit My Preference
        </button>
      ` : `
        <div style="margin-top:16px;text-align:center;color:#166534;font-weight:600;padding:12px;background:#dcfce7;border-radius:12px;">
          ✔ Thank you! We have already received your preferences.
        </div>
      `}

      <div id="statusMsg"></div>
    </form>

    <div class="footer">
      AutoVyn Insurance Services &bull; Secure Computer Generated Portal
    </div>
  </div>
</div>

<script>
function handleRadioSelect(radioElem) {
  const groupName = radioElem.name;
  const card = radioElem.closest('.question-card');
  document.querySelectorAll('input[name="' + groupName + '"]').forEach(r => {
    r.parentElement.classList.remove('selected');
  });
  if (radioElem.checked) {
    radioElem.parentElement.classList.add('selected');
  }

  // Toggle Other text input
  const otherInput = card.querySelector('.other-text-input');
  if (otherInput) {
    const isOtherSelected = radioElem.value.toLowerCase().startsWith('other');
    otherInput.style.display = isOtherSelected ? 'block' : 'none';
    if (isOtherSelected) otherInput.focus();
  }
}

function handleCheckboxSelect(cbElem) {
  cbElem.parentElement.classList.toggle('selected', cbElem.checked);
  const card = cbElem.closest('.question-card');
  const otherInput = card.querySelector('.other-text-input');
  if (otherInput && cbElem.value.toLowerCase().startsWith('other')) {
    otherInput.style.display = cbElem.checked ? 'block' : 'none';
    if (cbElem.checked) otherInput.focus();
  }
}

function submitPreferences() {
  const submitBtn = document.getElementById('submitBtn');
  const msg = document.getElementById('statusMsg');

  // Collect answers
  const questionCards = document.querySelectorAll('.question-card');
  const answers = [];

  for (const card of questionCards) {
    const qid = card.getAttribute('data-qid');
    const qtype = card.getAttribute('data-qtype');
    const qtext = card.getAttribute('data-qtext');
    let answerVal = '';

    if (qtype === 'text') {
      const input = card.querySelector('input[type="text"]');
      answerVal = input ? input.value.trim() : '';
    } else if (qtype === 'checkbox') {
      const checkedBoxes = card.querySelectorAll('input[type="checkbox"]:checked');
      const vals = [];
      checkedBoxes.forEach(c => {
        if (c.value.toLowerCase().startsWith('other')) {
          const otherText = card.querySelector('.other-text-input')?.value?.trim();
          vals.push(otherText ? ('Other: ' + otherText) : 'Other');
        } else {
          vals.push(c.value);
        }
      });
      answerVal = vals.join(', ');
    } else {
      const selectedRadio = card.querySelector('input[type="radio"]:checked');
      if (selectedRadio) {
        if (selectedRadio.value.toLowerCase().startsWith('other')) {
          const otherText = card.querySelector('.other-text-input')?.value?.trim();
          answerVal = otherText ? ('Other: ' + otherText) : 'Other';
        } else {
          answerVal = selectedRadio.value;
        }
      }
    }

    answers.push({
      question_id: parseInt(qid, 10),
      question_text: qtext,
      answer_text: answerVal
    });
  }

  // Disable button and show loader
  submitBtn.disabled = true;
  msg.className = 'loading';
  msg.textContent = '⏳ Saving your preferences...';

  fetch('${process.env.BASE_URL || ''}/Crm/SaveInsuranceCustomerResponse', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'compcode': '${rawCompCode}'
    },
    body: JSON.stringify({
      compcode: '${rawCompCode}',
      utd: '${rawUtd}',
      answers: answers
    })
  })
  .then(res => res.json())
  .then(data => {
    if (data.success) {
      msg.className = 'success';
      msg.textContent = '✅ ' + (data.message || 'Preferences submitted successfully!');
      setTimeout(() => { window.location.reload(); }, 2000);
    } else {
      msg.className = 'error';
      msg.textContent = '❌ ' + (data.message || 'Something went wrong. Please try again.');
      submitBtn.disabled = false;
    }
  })
  .catch(err => {
    msg.className = 'error';
    msg.textContent = '❌ Network error. Please try again.';
    submitBtn.disabled = false;
  });
}
</script>

</body>
</html>
    `);

  } catch (err) {
    console.error("InsuranceRenewalCustomerView ERROR:", err);
    res.send(`
      <!DOCTYPE html>
      <html>
      <head><title>Error</title><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
      <body style="font-family:Segoe UI, sans-serif;text-align:center;padding:50px;">
        <h2 style="color:#dc2626;">Error Loading Page</h2>
        <p style="color:#4b5563;">${err.message}</p>
      </body>
      </html>
    `);
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};


// ============================================================
// SAVE CUSTOMER RESPONSES INTO INSU_ANSWERS TABLE
// ============================================================
exports.SaveInsuranceCustomerResponse = async function (req, res) {
  let sequelize;
  try {
    const { compcode, utd, answers, cust_remark } = req.body || {};
    const effectiveCompCode = compcode || req.headers.compcode;

    if (!effectiveCompCode || !utd) {
      return res.status(400).send({
        success: false,
        message: "compcode and utd are required",
      });
    }

    sequelize = await dbname(req, effectiveCompCode);

    // 1. Fetch Customer Name from INSU_RENEWAL by Master UTD
    const [[custRow]] = await sequelize.query(`
      SELECT TOP 1 r.CUST_NAME
      FROM dbo.INSU_RENEWAL r
      WHERE r.TRAN_ID = '${utd}' AND r.EXPORT_TYPE = 1
      ORDER BY r.UTD DESC
    `);
    const customerName = String(custRow?.CUST_NAME || 'CUSTOMER').trim().replace(/'/g, "''").substring(0, 50);

    // 2. Archive previous active answers for this Customer / Renewal
    await sequelize.query(`
      UPDATE dbo.INSU_ANSWERS
      SET EXPORT_TYPE = 33
      WHERE Created_By = '${customerName}' AND (EXPORT_TYPE = 1 OR EXPORT_TYPE IS NULL)
    `);

    // 3. Insert answers: TRAN_ID = Question UTD (from INSU_QUESTIONS), Created_By = Customer Name
    if (Array.isArray(answers) && answers.length > 0) {
      for (const ans of answers) {
        const qId = ans.question_id || null;
        const aText = ans.answer_text ? String(ans.answer_text).trim() : '';
        const cleanAnswerText = aText.replace(/'/g, "''");

        if (aText && qId) {
          await sequelize.query(`
            INSERT INTO dbo.INSU_ANSWERS (
              Answer_Text, TRAN_ID, EXPORT_TYPE, Created_By, CREATED_AT
            ) VALUES (
              '${cleanAnswerText}', 
              '${qId}', 
              1, 
              '${customerName}', 
              GETDATE()
            )
          `);
        }
      }
    }

    return res.status(200).send({
      success: true,
      message: "Your preferences have been recorded successfully!",
    });

  } catch (e) {
    console.error("SaveInsuranceCustomerResponse ERROR:", e);
    return res.status(500).send({ success: false, message: e.message });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};


function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================================
// GET CUSTOMER INSURANCE QUESTIONS & ANSWERS RESPONSES (FOR MODAL)
// ============================================================
exports.getCustomerInsuranceResponses = async function (req, res) {
  let sequelize;
  try {
    const { cust_name, vehicle_number, vehical_reg_no, mob_no, utd } = req.body || {};
    const compcode = req.headers.compcode || req.body?.compcode;

    if (!compcode) {
      return res.status(400).send({ success: false, message: "compcode is required" });
    }

    sequelize = await dbname(req, compcode);

    let customerName = cust_name ? String(cust_name).trim() : "";
    let vehicleNo = vehicle_number || vehical_reg_no || "";
    let callId = null;

    // If customerName is not passed or we have UTD/vehicle, resolve customer name and get CALL_ID
    if (!customerName && (utd || vehicleNo || mob_no)) {
      const whereConditions = [];
      const replacements = {};

      if (utd) {
        whereConditions.push(`(m.UTD = :utd OR r.TRAN_ID = :utd)`);
        replacements.utd = utd;
      }
      if (vehicleNo) {
        whereConditions.push(`(m.VEHICAL_REG_NO = :vehNo OR r.VEHICAL_REG_NO = :vehNo)`);
        replacements.vehNo = String(vehicleNo).trim().toUpperCase();
      }
      if (mob_no) {
        whereConditions.push(`(r.CUST_MOB_NO LIKE :mob)`);
        replacements.mob = `%${String(mob_no).replace(/\D/g, "").slice(-10)}%`;
      }

      if (whereConditions.length > 0) {
        const [[custRow]] = await sequelize.query(`
          SELECT TOP 1 r.CUST_NAME, m.VEHICAL_REG_NO, r.TRAN_ID, r.CUST_MOB_NO
          FROM dbo.INSU_RENEWAL r
          LEFT JOIN dbo.INSU_RENEWAL_MST m ON m.UTD = r.TRAN_ID
          WHERE (${whereConditions.join(" OR ")})
            AND (r.EXPORT_TYPE = 1 OR r.EXPORT_TYPE IS NULL)
          ORDER BY r.UTD DESC
        `, { replacements });

        if (custRow) {
          customerName = custRow.CUST_NAME || "";
          vehicleNo = custRow.VEHICAL_REG_NO || vehicleNo;

          // ✅ NOW: Get CALL_ID from call_webhook_dtl table
          // Match करो: CUST_MOB_NO या TRAN_ID के हिसाब से
          let callIdQuery = `
            SELECT TOP 1 CALL_ID
            FROM dbo.call_webhook_dtl
            WHERE 1=1
          `;

          if (custRow.CUST_MOB_NO) {
            callIdQuery += ` AND CUST_MOB_NO = :mobNo`;
            replacements.mobNo = custRow.CUST_MOB_NO;
          }

          callIdQuery += ` ORDER BY CREATED_AT DESC`;

          const [[callRow]] = await sequelize.query(callIdQuery, { replacements });
          if (callRow && callRow.CALL_ID) {
            callId = callRow.CALL_ID;
          }
        }
      }
    }

    const cleanCustName = customerName ? customerName.replace(/'/g, "''").substring(0, 50) : "";

    // Fetch Questions and matching Answers filtered by CALL_ID if available
    let questionsQuery = `
      SELECT 
        q.UTD AS question_id,
        q.Question_Text,
        q.Question_Type,
        q.PAYMENT_AMOUNT,
        q.Options,
        q.Sort_Order,
        a.UTD AS answer_id,
        a.Answer_Text,
        CONVERT(varchar(19), a.CREATED_AT, 120) AS answered_at,
        a.Created_By AS answered_by,
        a.CALL_ID
      FROM dbo.INSU_QUESTIONS q
      LEFT JOIN dbo.INSU_ANSWERS a 
        ON CAST(a.TRAN_ID AS VARCHAR) = CAST(q.UTD AS VARCHAR)
        AND a.Created_By = :cleanCustName
        AND (a.EXPORT_TYPE = 1 OR a.EXPORT_TYPE IS NULL)
    `;

    const replacements2 = { cleanCustName };

    // ✅ Add CALL_ID filter if available
    if (callId) {
      questionsQuery += ` AND a.CALL_ID = :callId`;
      replacements2.callId = callId;
    }

    questionsQuery += `
      WHERE (q.Is_Active = '1' OR q.Is_Active = 'true' OR q.Is_Active = 'Y' OR q.Is_Active IS NULL)
      ORDER BY q.Sort_Order ASC, q.UTD ASC
    `;

    const questionsAndAnswers = await sequelize.query(questionsQuery, {
      replacements: replacements2,
      type: sequelize.QueryTypes.SELECT,
    });

    const hasSubmitted = (questionsAndAnswers || []).some(
      (qa) => qa.Answer_Text && String(qa.Answer_Text).trim().length > 0
    );

    const firstAnswer = (questionsAndAnswers || []).find(
      (qa) => qa.Answer_Text && qa.answered_at
    );

    return res.status(200).send({
      success: true,
      customerName,
      vehicleNo,
      callId,
      hasSubmitted,
      submittedAt: firstAnswer ? firstAnswer.answered_at : null,
      responses: questionsAndAnswers || [],
    });

  } catch (err) {
    console.error("getCustomerInsuranceResponses ERROR:", err);
    return res.status(500).send({ success: false, message: err.message });
  } finally {
    if (sequelize) {
      try {
        await sequelize.close();
      } catch (_) { }
    }
  }
};