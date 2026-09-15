const { Sequelize, DataTypes, Op } = require("sequelize");
const { dbname } = require("../utils/dbconfig");
const { v4: uuidv4 } = require("uuid")

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * HELPER: Build Manager Scope CTE / Subquery & Parameter Replacements
 * ══════════════════════════════════════════════════════════════════════════════
 * Resolves the authorized active employee scope for the manager.
 * Respects:
 * 1. Approval_Matrix (Direct approver 1/2/3)
 * 2. Department_Branch_Approval_Matrix (Dept + Branch approver)
 * 3. user_tbl.Multi_loc (Location permissions)
 * 4. EmployeeMaster.SUPERVISOR / SUPERVISORID (Direct reports)
 * 5. Frontend filters (applied strictly WITHIN manager scope)
 */
function buildManagerScopeQuery(params) {
  const {
    managerEmpCode,
    Loc_code,
    dept_code,
    shiftstarttime,
    selectedDate
  } = params;

  let extraEmpFilters = "";
  const replacements = {
    selectedDate: selectedDate,
    managerEmpCode: managerEmpCode || ""
  };

  // Safe Location Filter
  if (Loc_code !== undefined && Loc_code !== null && Loc_code !== "") {
    if (Array.isArray(Loc_code) && Loc_code.length > 0) {
      const locParams = Loc_code.map((_, idx) => `:loc_${idx}`);
      Loc_code.forEach((val, idx) => {
        replacements[`loc_${idx}`] = String(val);
      });
      extraEmpFilters += ` AND (EM.LOCATION IN (${locParams.join(",")}) OR CAST(EM.Loc_Code AS VARCHAR(20)) IN (${locParams.join(",")}))`;
    } else if (typeof Loc_code === "string" || typeof Loc_code === "number") {
      replacements.locSingle = String(Loc_code);
      extraEmpFilters += ` AND (EM.LOCATION = :locSingle OR CAST(EM.Loc_Code AS VARCHAR(20)) = :locSingle)`;
    }
  }

  // Safe Department Filter
  if (dept_code !== undefined && dept_code !== null && dept_code !== "") {
    replacements.deptCode = String(dept_code);
    extraEmpFilters += ` AND EM.Division = :deptCode`;
  }

  // Safe Shift Filter
  let shiftFilter = "";
  if (shiftstarttime !== undefined && shiftstarttime !== null && shiftstarttime !== "") {
    replacements.shiftStartTime = String(shiftstarttime);
    shiftFilter = ` AND AT.SHIFTSTARTTIME = :shiftStartTime`;
  }

  // Manager Scope Clause:
  // If managerEmpCode is provided, resolve manager's scope. If empty/admin, scope covers all active employees.
  let managerScopeCondition = "";
  if (managerEmpCode) {
    managerScopeCondition = `
      AND (
        -- 1. Direct Matrix Approval
        EM.EmpCode IN (
          SELECT empcode FROM Approval_Matrix WITH (NOLOCK)
          WHERE (approver1_A = :managerEmpCode OR approver1_B = :managerEmpCode
                 OR approver2_A = :managerEmpCode OR approver2_B = :managerEmpCode
                 OR approver3_A = :managerEmpCode OR approver3_B = :managerEmpCode)
        )
        -- 2. Department & Branch Approval Matrix
        OR EXISTS (
          SELECT 1 FROM Department_Branch_Approval_Matrix DBAM WITH (NOLOCK)
          WHERE (DBAM.approver1_A = :managerEmpCode OR DBAM.approver1_B = :managerEmpCode OR DBAM.approver1_C = :managerEmpCode
                 OR DBAM.approver2_A = :managerEmpCode OR DBAM.approver2_B = :managerEmpCode OR DBAM.approver2_C = :managerEmpCode
                 OR DBAM.approver3_A = :managerEmpCode OR DBAM.approver3_B = :managerEmpCode OR DBAM.approver3_C = :managerEmpCode)
            AND DBAM.department = EM.Division
            AND (DBAM.branch = EM.LOCATION OR DBAM.branch = CAST(EM.Loc_Code AS VARCHAR(20)))
        )
        -- 3. Direct Reporting / Supervisor
        OR EM.SUPERVISOR = :managerEmpCode
        OR CAST(EM.SUPERVISORID AS VARCHAR(50)) = :managerEmpCode
        -- 4. User Multi_loc Location Rights (if manager is branch/location head)
        OR (
          EXISTS (
            SELECT 1 FROM user_tbl UT WITH (NOLOCK)
            CROSS APPLY (
              SELECT CAST('<x>' + REPLACE(ISNULL(UT.Multi_loc, ''), ',', '</x><x>') + '</x>' AS XML)
            ) AS t(x)
            CROSS APPLY t.x.nodes('/x') AS a(m)
            WHERE UT.empcode = :managerEmpCode
              AND UT.export_type < 3
              AND (
                LTRIM(RTRIM(m.value('.', 'VARCHAR(50)'))) = EM.LOCATION
                OR LTRIM(RTRIM(m.value('.', 'VARCHAR(50)'))) = CAST(EM.Loc_Code AS VARCHAR(20))
              )
          )
        )
      )
    `;
  }

  const authorizedEmpCTE = `
    WITH AuthorizedEmployees AS (
      SELECT 
        EM.EmpCode,
        EM.EmpFirstName,
        EM.EmpLastName,
        CONCAT(ISNULL(EM.Title, ''), CASE WHEN EM.Title IS NOT NULL THEN ' ' ELSE '' END, EM.EmpFirstName, ' ', ISNULL(EM.EmpLastName, '')) AS FULL_NAME,
        EM.EMPLOYEEDesignation AS DESIGNATION,
        EM.EmployeeType AS EMPLOYEE_TYPE,
        EM.Division AS DEPT_CODE,
        EM.LOCATION AS LOC_CODE_STR,
        EM.Loc_Code AS LOC_CODE_NUM,
        EM.MOBILE_NO,
        EM.CORPORATEMAILID,
        EM.CURRENTJOINDATE
      FROM EMPLOYEEMASTER EM WITH (NOLOCK)
      WHERE EM.LASTWOR_DATE IS NULL
        ${managerScopeCondition}
        ${extraEmpFilters}
    )
  `;

  return {
    authorizedEmpCTE,
    replacements,
    shiftFilter,
    extraEmpFilters
  };
}

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * 1. MAIN ENDPOINT: GET MANAGER DASHBOARD DATA (Combined Read-Model)
 * ══════════════════════════════════════════════════════════════════════════════
 */
exports.getManagerDashboardData = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  try {
    const managerEmpCode = req.body.empcode || req.body.manager_empcode || req.body.EMPCODE || req.headers.empcode || "";
    const { Loc_code, dept_code, shiftstarttime } = req.body;

    let date = req.body.selectedDate || req.body.date;
    if (!date) {
      date = new Date().toISOString().split("T")[0];
    }

    const selectedDateObj = new Date(date);
    const selectedMonth = req.body.month || (selectedDateObj.getMonth() + 1);
    const selectedYear = req.body.year || selectedDateObj.getFullYear();

    const { authorizedEmpCTE, replacements, shiftFilter } = buildManagerScopeQuery({
      managerEmpCode,
      Loc_code,
      dept_code,
      shiftstarttime,
      selectedDate: date
    });

    replacements.selectedMonth = selectedMonth;
    replacements.selectedYear = selectedYear;

    const shiftStartExpr = `
      DATEADD(
        MINUTE,
        (FLOOR(AT.SHIFTSTARTTIME) * 60) + ROUND((AT.SHIFTSTARTTIME - FLOOR(AT.SHIFTSTARTTIME)) * 60, 0),
        CAST(AT.DATEOFFICE AS DATETIME)
      )
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 1: SHIFTS DROPDOWN
    // ──────────────────────────────────────────────────────────────────────────
    const shiftQuery = `
      ${authorizedEmpCTE}
      SELECT DISTINCT
        AT.SHIFTSTARTTIME,
        AT.SHIFTENDTIME,
        AT.SHIFT
      FROM attendancetable AT WITH (NOLOCK)
      INNER JOIN AuthorizedEmployees AE
        ON AT.Emp_Code = AE.EmpCode
      WHERE CAST(AT.DATEOFFICE AS DATE) = :selectedDate
      ORDER BY AT.SHIFTSTARTTIME
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 2: SUMMARY METRICS (Team Strength, Present Today, Mispunches, Waiting On Me)
    // ──────────────────────────────────────────────────────────────────────────
    const summaryQuery = `
      ${authorizedEmpCTE},
      SummaryMetrics AS (
        SELECT
          -- 1. Team Strength
          (SELECT COUNT(1) FROM AuthorizedEmployees) AS TEAM_STRENGTH,

          -- 2. Present Today (Deduplicated)
          (
            SELECT COUNT(DISTINCT AT.Emp_Code)
            FROM attendancetable AT WITH (NOLOCK)
            INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
            WHERE CAST(AT.DATEOFFICE AS DATE) = :selectedDate
              AND (
                AT.FLAG = 'P'
                OR (AT.IN1 IS NOT NULL AND ISNULL(AT.STATUS, '') <> 'A')
                OR (AT.STATUS IN ('P', 'PR', 'P/A', 'HD') OR ISNULL(AT.presentvalue, 0) > 0)
              )
              ${shiftFilter}
          ) AS PRESENT_TODAY,

          -- 2b. On Leave Today
          (
            SELECT COUNT(DISTINCT AT.Emp_Code)
            FROM attendancetable AT WITH (NOLOCK)
            INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
            WHERE CAST(AT.DATEOFFICE AS DATE) = :selectedDate
              AND (AT.LEAVEVALUE > 0 OR AT.STATUS IN ('L', 'HD') OR AT.Leave_applied_on IS NOT NULL)
              ${shiftFilter}
          ) AS ON_LEAVE_TODAY,

          -- 3. Monthly Mispunches Count
          (
            SELECT COUNT(1)
            FROM attendancetable AT WITH (NOLOCK)
            INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
            WHERE (
              (MONTH(AT.DATEOFFICE) = :selectedMonth AND YEAR(AT.DATEOFFICE) = :selectedYear)
              OR (MONTH(AT.Mispunch_applied_on) = :selectedMonth AND YEAR(AT.Mispunch_applied_on) = :selectedYear)
              OR (MONTH(AT.Mispunch_out_applied_on) = :selectedMonth AND YEAR(AT.Mispunch_out_applied_on) = :selectedYear)
            )
            AND (AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL)
          ) AS MONTHLY_MISPUNCHES,

          -- 4. Waiting on Me (Pending Approvals where manager is current approver)
          (
            SELECT COUNT(1)
            FROM attendancetable AT WITH (NOLOCK)
            INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
            WHERE (
              -- Leave Pending
              (AT.Leave_applied_on IS NOT NULL AND (
                (AT.Appr_1_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_1_Code = :managerEmpCode))
                OR (AT.Appr_1_Stat = 1 AND AT.Appr_2_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_2_Code = :managerEmpCode))
                OR (AT.Appr_2_Stat = 1 AND AT.Appr_3_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_3_Code = :managerEmpCode))
              ))
              -- Mispunch Pending
              OR ((AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL) AND (
                (AT.Appr_1_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_1_Code = :managerEmpCode))
                OR (AT.Appr_1_Stat = 1 AND AT.Appr_2_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_2_Code = :managerEmpCode))
                OR (AT.Appr_2_Stat = 1 AND AT.Appr_3_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_3_Code = :managerEmpCode))
              ))
            )
            AND ISNULL(AT.MAN_APPR, 'N') <> 'Y'
            AND ISNULL(AT.MAN_REJ, 'N') <> 'Y'
          ) AS WAITING_ON_ME
      )
      SELECT * FROM SummaryMetrics;
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 3: LAST 7 DAYS ATTENDANCE (Daily Breakdown: On-Time, Late, Leave, Absent)
    // ──────────────────────────────────────────────────────────────────────────
    const last7DaysQuery = `
      ${authorizedEmpCTE},
      DateRange AS (
        SELECT CAST(:selectedDate AS DATE) AS dt
        UNION ALL
        SELECT DATEADD(DAY, -1, dt) FROM DateRange WHERE dt > DATEADD(DAY, -7, :selectedDate)
      ),
      AT_DEDUP AS (
        SELECT
          AT.Emp_Code,
          CAST(AT.DATEOFFICE AS DATE) AS DATEOFFICE,
          AT.FLAG,
          AT.STATUS,
          AT.IN1,
          AT.OUT1,
          AT.LEAVEVALUE,
          AT.Leave_applied_on,
          ${shiftStartExpr} AS SHIFT_START_EXACT,
          ROW_NUMBER() OVER (
            PARTITION BY AT.Emp_Code, CAST(AT.DATEOFFICE AS DATE)
            ORDER BY CASE WHEN AT.IN1 IS NOT NULL THEN 0 ELSE 1 END, AT.IN1 DESC
          ) AS rn
        FROM attendancetable AT WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
        WHERE CAST(AT.DATEOFFICE AS DATE) BETWEEN DATEADD(DAY, -7, :selectedDate) AND :selectedDate
        ${shiftFilter}
      ),
      DAILY_STATUS AS (
        SELECT
          DR.dt AS DATE_LABEL,
          DATENAME(WEEKDAY, DR.dt) AS DAY_NAME,
          AE.EmpCode,
          CASE
            -- 1. Leave
            WHEN AT.LEAVEVALUE > 0 OR AT.STATUS IN ('L', 'HD') OR (AT.Leave_applied_on IS NOT NULL AND AT.STATUS = 'L')
              THEN 'LEAVE'
            -- 2. Absent (No punch or explicit A)
            WHEN AT.Emp_Code IS NULL
                 OR (DR.dt = CAST(GETDATE() AS DATE) AND (AT.FLAG = 'A' OR AT.FLAG IS NULL))
                 OR (DR.dt <> CAST(GETDATE() AS DATE) AND (AT.STATUS = 'A' OR AT.IN1 IS NULL))
              THEN 'ABSENT'
            -- 3. On Time
            WHEN AT.IN1 IS NOT NULL AND AT.IN1 <= AT.SHIFT_START_EXACT
              THEN 'ON_TIME'
            -- 4. Late
            WHEN AT.IN1 IS NOT NULL AND AT.IN1 > AT.SHIFT_START_EXACT
              THEN 'LATE'
            -- Fallback Present
            ELSE 'ON_TIME'
          END AS ATTENDANCE_STATUS
        FROM DateRange DR
        CROSS JOIN AuthorizedEmployees AE
        LEFT JOIN AT_DEDUP AT
          ON AT.Emp_Code = AE.EmpCode
         AND AT.DATEOFFICE = DR.dt
         AND AT.rn = 1
      )
      SELECT
        DATE_LABEL,
        DAY_NAME,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'ON_TIME' THEN 1 END) AS ON_TIME,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'LATE' THEN 1 END) AS LATE,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'LEAVE' THEN 1 END) AS ON_LEAVE,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'ABSENT' THEN 1 END) AS ABSENT,
        COUNT(1) AS TOTAL_EMP,
        CASE
          WHEN COUNT(1) > 0 THEN
            ROUND((CAST(COUNT(CASE WHEN ATTENDANCE_STATUS IN ('ON_TIME', 'LATE') THEN 1 END) AS FLOAT) / CAST(COUNT(1) AS FLOAT)) * 100, 2)
          ELSE 0
        END AS ATTENDANCE_PCT
      FROM DAILY_STATUS
      GROUP BY DATE_LABEL, DAY_NAME
      ORDER BY DATE_LABEL ASC
      OPTION (MAXRECURSION 10);
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 4: DEPARTMENT-WISE ATTENDANCE STATS
    // ──────────────────────────────────────────────────────────────────────────
    const deptAttendanceQuery = `
      ${authorizedEmpCTE},
      DeptAttendance AS (
        SELECT
          ISNULL(MM.Misc_Code, AE.DEPT_CODE) AS DEPT_CODE,
          ISNULL(MM.Misc_Name, 'Unassigned Department') AS DEPT_NAME,
          COUNT(AE.EmpCode) AS TOTAL,
          COUNT(
            CASE
              WHEN CAST(AT.DATEOFFICE AS DATE) = :selectedDate AND AT.FLAG = 'P' THEN 1
              WHEN CAST(AT.DATEOFFICE AS DATE) <> :selectedDate AND AT.IN1 IS NOT NULL AND AT.STATUS <> 'A' THEN 1
            END
          ) AS PRESENT,
          COUNT(
            CASE
              WHEN CAST(AT.DATEOFFICE AS DATE) = :selectedDate AND (AT.FLAG = 'A' OR AT.FLAG IS NULL) THEN 1
              WHEN CAST(AT.DATEOFFICE AS DATE) <> :selectedDate AND (AT.IN1 IS NULL OR AT.STATUS = 'A') THEN 1
            END
          ) AS ABSENT,
          COUNT(
            CASE
              WHEN AT.IN1 IS NOT NULL AND AT.IN1 <= ${shiftStartExpr} THEN 1
            END
          ) AS ON_TIME,
          COUNT(
            CASE
              WHEN AT.IN1 IS NOT NULL AND AT.IN1 > ${shiftStartExpr} THEN 1
            END
          ) AS LATE,
          COUNT(
            CASE
              WHEN AT.LEAVEVALUE > 0 OR AT.STATUS IN ('L', 'HD') THEN 1
            END
          ) AS ON_LEAVE
        FROM AuthorizedEmployees AE
        LEFT JOIN MISC_MST MM WITH (NOLOCK)
          ON MM.Misc_Type = 68
         AND MM.Misc_Code = AE.DEPT_CODE
         AND ISNULL(MM.Export_Type, 0) < 3
        LEFT JOIN attendancetable AT WITH (NOLOCK)
          ON AT.Emp_Code = AE.EmpCode
         AND CAST(AT.DATEOFFICE AS DATE) = :selectedDate
         ${shiftFilter}
        GROUP BY ISNULL(MM.Misc_Code, AE.DEPT_CODE), ISNULL(MM.Misc_Name, 'Unassigned Department')
      )
      SELECT * FROM DeptAttendance ORDER BY DEPT_NAME;
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 5: SALARY & DEDUCTIONS (Overall, Dept-Wise, Location-Wise)
    // ──────────────────────────────────────────────────────────────────────────
    const salaryQuery = `
      ${authorizedEmpCTE},
      LatestSalary AS (
        SELECT
          SS.Emp_Code,
          SS.Gross_Salary,
          SS.Loc_Code,
          ROW_NUMBER() OVER (
            PARTITION BY SS.Emp_Code 
            ORDER BY ISNULL(SS.Effective_date, '1900-01-01') DESC, ISNULL(SS.Rec_date, '1900-01-01') DESC, SS.ServerId DESC
          ) AS rn
        FROM SALARYSTRUCTURE SS WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON SS.Emp_Code = AE.EmpCode
        WHERE ISNULL(SS.Export_Type, 0) < 3
      ),
      EmpDeductions AS (
        SELECT
          ED.Emp_Id,
          SUM(ISNULL(ED.Ded_Amt, 0)) AS TOTAL_DED
        FROM Emp_Ded ED WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON ED.Emp_Id = AE.EmpCode
        WHERE ISNULL(ED.Export_type, 0) < 3
          AND (:selectedMonth IS NULL OR ED.Mnth = :selectedMonth)
          AND (:selectedYear IS NULL OR ED.Yr = :selectedYear)
        GROUP BY ED.Emp_Id
      ),
      EmployeeCompensation AS (
        SELECT
          AE.EmpCode,
          AE.DEPT_CODE,
          ISNULL(MM_DEPT.Misc_Name, 'Unassigned Department') AS DEPT_NAME,
          ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM) AS SALARY_LOC_CODE,
          ISNULL(GM.Godw_Name, ISNULL(MM_LOC.Misc_Name, 'Unassigned Location')) AS SALARY_LOC_NAME,
          ISNULL(LS.Gross_Salary, 0) AS GROSS_SALARY,
          ISNULL(ED.TOTAL_DED, 0) AS DEDUCTION_AMT,
          (ISNULL(LS.Gross_Salary, 0) - ISNULL(ED.TOTAL_DED, 0)) AS NET_PAY
        FROM AuthorizedEmployees AE
        LEFT JOIN LatestSalary LS ON LS.Emp_Code = AE.EmpCode AND LS.rn = 1
        LEFT JOIN EmpDeductions ED ON ED.Emp_Id = AE.EmpCode
        LEFT JOIN MISC_MST MM_DEPT WITH (NOLOCK)
          ON MM_DEPT.Misc_Type = 68
         AND MM_DEPT.Misc_Code = AE.DEPT_CODE
         AND ISNULL(MM_DEPT.Export_Type, 0) < 3
        LEFT JOIN godown_mst GM WITH (NOLOCK)
          ON GM.Godw_Code = ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM)
         AND ISNULL(GM.export_type, 0) < 3
        LEFT JOIN MISC_MST MM_LOC WITH (NOLOCK)
          ON MM_LOC.Misc_Type = 85
         AND MM_LOC.Misc_Code = CAST(ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM) AS VARCHAR(20))
         AND ISNULL(MM_LOC.Export_Type, 0) < 3
      )
      SELECT
        -- Overall
        SUM(GROSS_SALARY) AS OVERALL_GROSS,
        SUM(DEDUCTION_AMT) AS OVERALL_DED,
        SUM(NET_PAY) AS OVERALL_NET,
        COUNT(1) AS TOTAL_EMPLOYEES
      FROM EmployeeCompensation;
    `;

    const salaryByDeptQuery = `
      ${authorizedEmpCTE},
      LatestSalary AS (
        SELECT
          SS.Emp_Code,
          SS.Gross_Salary,
          ROW_NUMBER() OVER (
            PARTITION BY SS.Emp_Code 
            ORDER BY ISNULL(SS.Effective_date, '1900-01-01') DESC, ISNULL(SS.Rec_date, '1900-01-01') DESC, SS.ServerId DESC
          ) AS rn
        FROM SALARYSTRUCTURE SS WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON SS.Emp_Code = AE.EmpCode
        WHERE ISNULL(SS.Export_Type, 0) < 3
      ),
      EmpDeductions AS (
        SELECT
          ED.Emp_Id,
          SUM(ISNULL(ED.Ded_Amt, 0)) AS TOTAL_DED
        FROM Emp_Ded ED WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON ED.Emp_Id = AE.EmpCode
        WHERE ISNULL(ED.Export_type, 0) < 3
          AND (:selectedMonth IS NULL OR ED.Mnth = :selectedMonth)
          AND (:selectedYear IS NULL OR ED.Yr = :selectedYear)
        GROUP BY ED.Emp_Id
      ),
      DeptComp AS (
        SELECT
          AE.DEPT_CODE,
          ISNULL(MM_DEPT.Misc_Name, 'Unassigned Department') AS DEPT_NAME,
          COUNT(AE.EmpCode) AS EMP_COUNT,
          SUM(ISNULL(LS.Gross_Salary, 0)) AS GROSS_SALARY,
          SUM(ISNULL(ED.TOTAL_DED, 0)) AS DEDUCTION_AMT,
          SUM(ISNULL(LS.Gross_Salary, 0) - ISNULL(ED.TOTAL_DED, 0)) AS NET_PAY
        FROM AuthorizedEmployees AE
        LEFT JOIN LatestSalary LS ON LS.Emp_Code = AE.EmpCode AND LS.rn = 1
        LEFT JOIN EmpDeductions ED ON ED.Emp_Id = AE.EmpCode
        LEFT JOIN MISC_MST MM_DEPT WITH (NOLOCK)
          ON MM_DEPT.Misc_Type = 68
         AND MM_DEPT.Misc_Code = AE.DEPT_CODE
         AND ISNULL(MM_DEPT.Export_Type, 0) < 3
        GROUP BY AE.DEPT_CODE, MM_DEPT.Misc_Name
      )
      SELECT * FROM DeptComp ORDER BY GROSS_SALARY DESC;
    `;

    const salaryByLocQuery = `
      ${authorizedEmpCTE},
      LatestSalary AS (
        SELECT
          SS.Emp_Code,
          SS.Gross_Salary,
          SS.Loc_Code,
          ROW_NUMBER() OVER (
            PARTITION BY SS.Emp_Code 
            ORDER BY ISNULL(SS.Effective_date, '1900-01-01') DESC, ISNULL(SS.Rec_date, '1900-01-01') DESC, SS.ServerId DESC
          ) AS rn
        FROM SALARYSTRUCTURE SS WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON SS.Emp_Code = AE.EmpCode
        WHERE ISNULL(SS.Export_Type, 0) < 3
      ),
      EmpDeductions AS (
        SELECT
          ED.Emp_Id,
          SUM(ISNULL(ED.Ded_Amt, 0)) AS TOTAL_DED
        FROM Emp_Ded ED WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON ED.Emp_Id = AE.EmpCode
        WHERE ISNULL(ED.Export_type, 0) < 3
          AND (:selectedMonth IS NULL OR ED.Mnth = :selectedMonth)
          AND (:selectedYear IS NULL OR ED.Yr = :selectedYear)
        GROUP BY ED.Emp_Id
      ),
      LocComp AS (
        SELECT
          ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM) AS LOC_CODE,
          ISNULL(GM.Godw_Name, ISNULL(MM_LOC.Misc_Name, 'Unassigned Location')) AS LOC_NAME,
          COUNT(AE.EmpCode) AS EMP_COUNT,
          SUM(ISNULL(LS.Gross_Salary, 0)) AS GROSS_SALARY,
          SUM(ISNULL(ED.TOTAL_DED, 0)) AS DEDUCTION_AMT,
          SUM(ISNULL(LS.Gross_Salary, 0) - ISNULL(ED.TOTAL_DED, 0)) AS NET_PAY
        FROM AuthorizedEmployees AE
        LEFT JOIN LatestSalary LS ON LS.Emp_Code = AE.EmpCode AND LS.rn = 1
        LEFT JOIN EmpDeductions ED ON ED.Emp_Id = AE.EmpCode
        LEFT JOIN godown_mst GM WITH (NOLOCK)
          ON GM.Godw_Code = ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM)
         AND ISNULL(GM.export_type, 0) < 3
        LEFT JOIN MISC_MST MM_LOC WITH (NOLOCK)
          ON MM_LOC.Misc_Type = 85
         AND MM_LOC.Misc_Code = CAST(ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM) AS VARCHAR(20))
         AND ISNULL(MM_LOC.Export_Type, 0) < 3
        GROUP BY ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM), GM.Godw_Name, MM_LOC.Misc_Name
      )
      SELECT * FROM LocComp ORDER BY GROSS_SALARY DESC;
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 6: MANAGER ACTIONS / WAITING ON ME DETAILS
    // ──────────────────────────────────────────────────────────────────────────
    const actionsQuery = `
      ${authorizedEmpCTE}
      SELECT
        AT.UTD AS REFERENCE_ID,
        AE.EmpCode AS EMP_CODE,
        AE.FULL_NAME AS EMP_NAME,
        AE.DESIGNATION,
        AE.DEPT_CODE,
        ISNULL(DEPT.Misc_Name, 'Department') AS DEPARTMENT,
        CAST(AT.DATEOFFICE AS DATE) AS ACTION_DATE,
        CASE
          WHEN AT.Leave_applied_on IS NOT NULL THEN 'LEAVE_APPROVAL'
          WHEN AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL THEN 'MISPUNCH_APPROVAL'
          ELSE 'GENERAL_APPROVAL'
        END AS ACTION_TYPE,
        CASE
          WHEN AT.Leave_applied_on IS NOT NULL THEN CONCAT('Leave Request: ', ISNULL(LM.Misc_Name, 'Leave'))
          WHEN AT.Mispunch_applied_on IS NOT NULL AND AT.Mispunch_out_applied_on IS NOT NULL THEN 'Mispunch In & Out Request'
          WHEN AT.Mispunch_applied_on IS NOT NULL THEN 'Mispunch In Request'
          WHEN AT.Mispunch_out_applied_on IS NOT NULL THEN 'Mispunch Out Request'
          ELSE 'Pending Approval'
        END AS TITLE,
        ISNULL(AT.SPL_REMARK, AT.reason) AS REMARKS,
        ISNULL(AT.Leave_applied_on, ISNULL(AT.Mispunch_applied_on, AT.Mispunch_out_applied_on)) AS APPLIED_ON,
        CASE
          WHEN AT.Appr_1_Stat IS NULL THEN 'Level 1 Pending'
          WHEN AT.Appr_2_Stat IS NULL THEN 'Level 2 Pending'
          WHEN AT.Appr_3_Stat IS NULL THEN 'Level 3 Pending'
          ELSE 'Pending'
        END AS PENDING_LEVEL
      FROM attendancetable AT WITH (NOLOCK)
      INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK)
        ON DEPT.Misc_Type = 68
       AND DEPT.Misc_Code = AE.DEPT_CODE
       AND ISNULL(DEPT.Export_Type, 0) < 3
      LEFT JOIN MISC_MST LM WITH (NOLOCK)
        ON LM.Misc_Type = 92
       AND LM.Misc_Code = AT.mipunch_reason
       AND ISNULL(LM.Export_Type, 0) < 3
      WHERE (
        -- Pending Leave
        (AT.Leave_applied_on IS NOT NULL AND (
          (AT.Appr_1_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_1_Code = :managerEmpCode))
          OR (AT.Appr_1_Stat = 1 AND AT.Appr_2_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_2_Code = :managerEmpCode))
          OR (AT.Appr_2_Stat = 1 AND AT.Appr_3_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_3_Code = :managerEmpCode))
        ))
        -- Pending Mispunch
        OR ((AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL) AND (
          (AT.Appr_1_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_1_Code = :managerEmpCode))
          OR (AT.Appr_1_Stat = 1 AND AT.Appr_2_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_2_Code = :managerEmpCode))
          OR (AT.Appr_2_Stat = 1 AND AT.Appr_3_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_3_Code = :managerEmpCode))
        ))
      )
      AND ISNULL(AT.MAN_APPR, 'N') <> 'Y'
      AND ISNULL(AT.MAN_REJ, 'N') <> 'Y'
      ORDER BY APPLIED_ON DESC;
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 7: WATCHLIST & EXCEPTIONS
    // ──────────────────────────────────────────────────────────────────────────
    const watchlistQuery = `
      ${authorizedEmpCTE},
      MonthlyMispunchSummary AS (
        SELECT
          AT.Emp_Code,
          COUNT(1) AS MISPUNCH_COUNT
        FROM attendancetable AT WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
        WHERE (MONTH(AT.DATEOFFICE) = :selectedMonth AND YEAR(AT.DATEOFFICE) = :selectedYear)
          AND (AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL)
        GROUP BY AT.Emp_Code
        HAVING COUNT(1) >= 2
      ),
      LateSummary AS (
        SELECT
          AT.Emp_Code,
          COUNT(1) AS LATE_COUNT
        FROM attendancetable AT WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
        WHERE CAST(AT.DATEOFFICE AS DATE) BETWEEN DATEADD(DAY, -7, :selectedDate) AND :selectedDate
          AND AT.IN1 IS NOT NULL
          AND AT.IN1 > ${shiftStartExpr}
        GROUP BY AT.Emp_Code
        HAVING COUNT(1) >= 2
      )
      SELECT
        AE.EmpCode AS employeeId,
        AE.FULL_NAME AS employeeName,
        ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS department,
        ISNULL(GM.Godw_Name, AE.LOC_CODE_STR) AS location,
        'HIGH_MISPUNCH' AS issueType,
        CONCAT(MS.MISPUNCH_COUNT, ' mispunches recorded in current month') AS message,
        CAST(MS.MISPUNCH_COUNT AS VARCHAR(20)) AS value,
        'MEDIUM' AS severity,
        AE.EmpCode AS referenceId
      FROM MonthlyMispunchSummary MS
      INNER JOIN AuthorizedEmployees AE ON MS.Emp_Code = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK) ON DEPT.Misc_Type = 68 AND DEPT.Misc_Code = AE.DEPT_CODE
      LEFT JOIN godown_mst GM WITH (NOLOCK) ON GM.Godw_Code = AE.LOC_CODE_NUM

      UNION ALL

      SELECT
        AE.EmpCode AS employeeId,
        AE.FULL_NAME AS employeeName,
        ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS department,
        ISNULL(GM.Godw_Name, AE.LOC_CODE_STR) AS location,
        'FREQUENT_LATE' AS issueType,
        CONCAT(LS.LATE_COUNT, ' late arrivals in last 7 days') AS message,
        CAST(LS.LATE_COUNT AS VARCHAR(20)) AS value,
        'HIGH' AS severity,
        AE.EmpCode AS referenceId
      FROM LateSummary LS
      INNER JOIN AuthorizedEmployees AE ON LS.Emp_Code = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK) ON DEPT.Misc_Type = 68 AND DEPT.Misc_Code = AE.DEPT_CODE
      LEFT JOIN godown_mst GM WITH (NOLOCK) ON GM.Godw_Code = AE.LOC_CODE_NUM;
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // QUERY 8: EMPLOYEE LIST (Right Panel with Photos)
    // ──────────────────────────────────────────────────────────────────────────
    const employeeListQuery = `
      ${authorizedEmpCTE}
      SELECT
        AE.EmpCode AS EMP_CODE,
        AE.FULL_NAME AS EMP_NAME,
        AE.DESIGNATION,
        AE.EMPLOYEE_TYPE,
        ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS DEPARTMENT,
        CONVERT(VARCHAR(8), AT.IN1, 108) AS LOGIN_TIME,
        CONVERT(VARCHAR(8), AT.OUT1, 108) AS LOGOUT_TIME,
        AT.STATUS AS ATTENDANCE_STATUS,
        CASE
          WHEN AT.FLAG = 'P' OR (AT.IN1 IS NOT NULL AND ISNULL(AT.STATUS, '') <> 'A') OR AT.STATUS IN ('P', 'PR', 'P/A', 'HD') OR ISNULL(AT.presentvalue, 0) > 0 THEN 'PRESENT'
          WHEN AT.LEAVEVALUE > 0 OR AT.STATUS IN ('L', 'HD') OR AT.Leave_applied_on IS NOT NULL THEN 'LEAVE'
          ELSE 'ABSENT'
        END AS PUNCH_STATUS,
        ED.DOC_PATH AS PHOTO_PATH
      FROM AuthorizedEmployees AE
      LEFT JOIN attendancetable AT WITH (NOLOCK)
        ON AT.Emp_Code = AE.EmpCode
       AND CAST(AT.DATEOFFICE AS DATE) = :selectedDate
       ${shiftFilter}
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK)
        ON DEPT.Misc_Type = 68
       AND DEPT.Misc_Code = AE.DEPT_CODE
       AND ISNULL(DEPT.Export_Type, 0) < 3
      LEFT JOIN EMP_DOCS ED WITH (NOLOCK)
        ON ED.EMP_CODE = AE.EmpCode
       AND ED.columndoc_type = 'EMPLOYEE'
       AND ED.Seq_No = 1
       AND ISNULL(ED.export_type, 0) < 3
      ORDER BY AE.EmpFirstName;
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // RUN ALL QUERIES IN PARALLEL
    // ──────────────────────────────────────────────────────────────────────────
    const [
      [shifts],
      [summaryRows],
      [last7Days],
      [deptAttendance],
      [overallSalaryRows],
      [salaryByDept],
      [salaryByLoc],
      [actionItems],
      [watchlist],
      [employees]
    ] = await Promise.all([
      sequelize.query(shiftQuery, { replacements }),
      sequelize.query(summaryQuery, { replacements }),
      sequelize.query(last7DaysQuery, { replacements }),
      sequelize.query(deptAttendanceQuery, { replacements }),
      sequelize.query(salaryQuery, { replacements }),
      sequelize.query(salaryByDeptQuery, { replacements }),
      sequelize.query(salaryByLocQuery, { replacements }),
      sequelize.query(actionsQuery, { replacements }),
      sequelize.query(watchlistQuery, { replacements }),
      sequelize.query(employeeListQuery, { replacements })
    ]);

    // Format Summary
    const teamStrength = Number(summaryRows[0]?.TEAM_STRENGTH || 0);
    const presentToday = Number(summaryRows[0]?.PRESENT_TODAY || 0);
    const onLeaveToday = Number(summaryRows[0]?.ON_LEAVE_TODAY || 0);
    const absentToday = Math.max(0, teamStrength - presentToday - onLeaveToday);
    const presentPercentage = teamStrength > 0 ? Number(((presentToday / teamStrength) * 100).toFixed(1)) : 0;

    const summary = {
      teamStrength,
      presentToday,
      onLeaveToday,
      absentToday,
      presentPercentage,
      waitingOnMe: Number(summaryRows[0]?.WAITING_ON_ME || 0),
      monthlyMispunches: Number(summaryRows[0]?.MONTHLY_MISPUNCHES || 0)
    };

    // Format Salary
    const overallGross = Number(overallSalaryRows[0]?.OVERALL_GROSS || 0);
    const overallDed = Number(overallSalaryRows[0]?.OVERALL_DED || 0);
    const overallNet = Number(overallSalaryRows[0]?.OVERALL_NET || (overallGross - overallDed));

    const formattedDeptSalary = salaryByDept.map(item => {
      const gross = Number(item.GROSS_SALARY || 0);
      const ded = Number(item.DEDUCTION_AMT || 0);
      const net = Number(item.NET_PAY || (gross - ded));
      const grossPct = overallGross > 0 ? Number(((gross / overallGross) * 100).toFixed(2)) : 0;
      return {
        deptCode: item.DEPT_CODE,
        deptName: item.DEPT_NAME,
        empCount: Number(item.EMP_COUNT || 0),
        gross,
        deduction: ded,
        netPay: net,
        grossPercentage: grossPct
      };
    });

    const formattedLocSalary = salaryByLoc.map(item => {
      const gross = Number(item.GROSS_SALARY || 0);
      const ded = Number(item.DEDUCTION_AMT || 0);
      const net = Number(item.NET_PAY || (gross - ded));
      const grossPct = overallGross > 0 ? Number(((gross / overallGross) * 100).toFixed(2)) : 0;
      return {
        locCode: item.LOC_CODE,
        locName: item.LOC_NAME,
        empCount: Number(item.EMP_COUNT || 0),
        gross,
        deduction: ded,
        netPay: net,
        grossPercentage: grossPct
      };
    });

    // Format Employees
    const BASE_URL = process.env.DOC_BASE_URL || "";
    const employeeList = employees.map(emp => ({
      ...emp,
      PHOTO_URL: emp.PHOTO_PATH ? `${BASE_URL}${emp.PHOTO_PATH}` : null
    }));

    const employeesData = {
      all: employeeList,
      present: employeeList.filter(e => e.PUNCH_STATUS === "PRESENT"),
      absent: employeeList.filter(e => e.PUNCH_STATUS === "ABSENT")
    };

    // Final Structured Response
    return res.status(200).json({
      success: true,
      date: date,
      meta: {
        month: selectedMonth,
        year: selectedYear,
        scope: {
          managerEmpCode: managerEmpCode || "ALL",
          employeeCount: summary.teamStrength
        },
        asOf: new Date().toISOString()
      },
      summary,
      attendance: {
        shifts,
        last7Days,
        departments: deptAttendance,
        employees: employeesData
      },
      salary: {
        overall: {
          gross: overallGross,
          deduction: overallDed,
          netPay: overallNet
        },
        byDepartment: formattedDeptSalary,
        byLocation: formattedLocSalary
      },
      actions: actionItems,
      watchlist: watchlist
    });

  } catch (err) {
    console.error("❌ getManagerDashboardData Error:", err);
    return res.status(500).json({ success: false, error: err.message });
  } finally {
    if (sequelize) await sequelize.close();
  }
};

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * 2. MODULAR ENDPOINT: ATTENDANCE STATS
 * ══════════════════════════════════════════════════════════════════════════════
 */
exports.getManagerAttendanceStats = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  try {
    const managerEmpCode = req.body.empcode || req.body.manager_empcode || req.body.EMPCODE || req.headers.empcode || "";
    const { Loc_code, dept_code, shiftstarttime } = req.body;
    let date = req.body.selectedDate || req.body.date || new Date().toISOString().split("T")[0];

    const { authorizedEmpCTE, replacements, shiftFilter } = buildManagerScopeQuery({
      managerEmpCode,
      Loc_code,
      dept_code,
      shiftstarttime,
      selectedDate: date
    });

    const shiftStartExpr = `
      DATEADD(
        MINUTE,
        (FLOOR(AT.SHIFTSTARTTIME) * 60) + ROUND((AT.SHIFTSTARTTIME - FLOOR(AT.SHIFTSTARTTIME)) * 60, 0),
        CAST(AT.DATEOFFICE AS DATETIME)
      )
    `;

    const last7DaysQuery = `
      ${authorizedEmpCTE},
      DateRange AS (
        SELECT CAST(:selectedDate AS DATE) AS dt
        UNION ALL
        SELECT DATEADD(DAY, -1, dt) FROM DateRange WHERE dt > DATEADD(DAY, -7, :selectedDate)
      ),
      AT_DEDUP AS (
        SELECT
          AT.Emp_Code,
          CAST(AT.DATEOFFICE AS DATE) AS DATEOFFICE,
          AT.FLAG,
          AT.STATUS,
          AT.IN1,
          AT.LEAVEVALUE,
          AT.Leave_applied_on,
          ${shiftStartExpr} AS SHIFT_START_EXACT,
          ROW_NUMBER() OVER (
            PARTITION BY AT.Emp_Code, CAST(AT.DATEOFFICE AS DATE)
            ORDER BY CASE WHEN AT.IN1 IS NOT NULL THEN 0 ELSE 1 END, AT.IN1 DESC
          ) AS rn
        FROM attendancetable AT WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
        WHERE CAST(AT.DATEOFFICE AS DATE) BETWEEN DATEADD(DAY, -7, :selectedDate) AND :selectedDate
        ${shiftFilter}
      ),
      DAILY_STATUS AS (
        SELECT
          DR.dt AS DATE_LABEL,
          DATENAME(WEEKDAY, DR.dt) AS DAY_NAME,
          AE.EmpCode,
          CASE
            WHEN AT.LEAVEVALUE > 0 OR AT.STATUS IN ('L', 'HD') OR (AT.Leave_applied_on IS NOT NULL AND AT.STATUS = 'L') THEN 'LEAVE'
            WHEN AT.Emp_Code IS NULL OR (DR.dt = CAST(GETDATE() AS DATE) AND (AT.FLAG = 'A' OR AT.FLAG IS NULL)) OR (DR.dt <> CAST(GETDATE() AS DATE) AND (AT.STATUS = 'A' OR AT.IN1 IS NULL)) THEN 'ABSENT'
            WHEN AT.IN1 IS NOT NULL AND AT.IN1 <= AT.SHIFT_START_EXACT THEN 'ON_TIME'
            WHEN AT.IN1 IS NOT NULL AND AT.IN1 > AT.SHIFT_START_EXACT THEN 'LATE'
            ELSE 'ON_TIME'
          END AS ATTENDANCE_STATUS
        FROM DateRange DR
        CROSS JOIN AuthorizedEmployees AE
        LEFT JOIN AT_DEDUP AT ON AT.Emp_Code = AE.EmpCode AND AT.DATEOFFICE = DR.dt AND AT.rn = 1
      )
      SELECT
        DATE_LABEL,
        DAY_NAME,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'ON_TIME' THEN 1 END) AS ON_TIME,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'LATE' THEN 1 END) AS LATE,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'LEAVE' THEN 1 END) AS ON_LEAVE,
        COUNT(CASE WHEN ATTENDANCE_STATUS = 'ABSENT' THEN 1 END) AS ABSENT,
        COUNT(1) AS TOTAL_EMP,
        CASE
          WHEN COUNT(1) > 0 THEN ROUND((CAST(COUNT(CASE WHEN ATTENDANCE_STATUS IN ('ON_TIME', 'LATE') THEN 1 END) AS FLOAT) / CAST(COUNT(1) AS FLOAT)) * 100, 2)
          ELSE 0
        END AS ATTENDANCE_PCT
      FROM DAILY_STATUS
      GROUP BY DATE_LABEL, DAY_NAME
      ORDER BY DATE_LABEL ASC
      OPTION (MAXRECURSION 10);
    `;

    const [[last7Days]] = await Promise.all([
      sequelize.query(last7DaysQuery, { replacements })
    ]);

    return res.status(200).json({ success: true, date, result: last7Days });
  } catch (err) {
    console.error("❌ getManagerAttendanceStats Error:", err);
    return res.status(500).json({ success: false, error: err.message });
  } finally {
    if (sequelize) await sequelize.close();
  }
};

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * 3. MODULAR ENDPOINT: SALARY STATS
 * ══════════════════════════════════════════════════════════════════════════════
 */
exports.getManagerSalaryStats = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  try {
    const managerEmpCode = req.body.empcode || req.body.manager_empcode || req.body.EMPCODE || req.headers.empcode || "";
    const { Loc_code, dept_code } = req.body;
    let date = req.body.selectedDate || req.body.date || new Date().toISOString().split("T")[0];
    const selectedDateObj = new Date(date);
    const selectedMonth = req.body.month || (selectedDateObj.getMonth() + 1);
    const selectedYear = req.body.year || selectedDateObj.getFullYear();

    const { authorizedEmpCTE, replacements } = buildManagerScopeQuery({
      managerEmpCode,
      Loc_code,
      dept_code,
      selectedDate: date
    });
    replacements.selectedMonth = selectedMonth;
    replacements.selectedYear = selectedYear;

    const salaryQuery = `
      ${authorizedEmpCTE},
      LatestSalary AS (
        SELECT
          SS.Emp_Code,
          SS.Gross_Salary,
          SS.Loc_Code,
          ROW_NUMBER() OVER (
            PARTITION BY SS.Emp_Code 
            ORDER BY ISNULL(SS.Effective_date, '1900-01-01') DESC, ISNULL(SS.Rec_date, '1900-01-01') DESC, SS.ServerId DESC
          ) AS rn
        FROM SALARYSTRUCTURE SS WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON SS.Emp_Code = AE.EmpCode
        WHERE ISNULL(SS.Export_Type, 0) < 3
      ),
      EmpDeductions AS (
        SELECT
          ED.Emp_Id,
          SUM(ISNULL(ED.Ded_Amt, 0)) AS TOTAL_DED
        FROM Emp_Ded ED WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON ED.Emp_Id = AE.EmpCode
        WHERE ISNULL(ED.Export_type, 0) < 3
          AND (:selectedMonth IS NULL OR ED.Mnth = :selectedMonth)
          AND (:selectedYear IS NULL OR ED.Yr = :selectedYear)
        GROUP BY ED.Emp_Id
      ),
      EmployeeCompensation AS (
        SELECT
          AE.EmpCode,
          AE.DEPT_CODE,
          ISNULL(MM_DEPT.Misc_Name, 'Unassigned Department') AS DEPT_NAME,
          ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM) AS SALARY_LOC_CODE,
          ISNULL(GM.Godw_Name, ISNULL(MM_LOC.Misc_Name, 'Unassigned Location')) AS SALARY_LOC_NAME,
          ISNULL(LS.Gross_Salary, 0) AS GROSS_SALARY,
          ISNULL(ED.TOTAL_DED, 0) AS DEDUCTION_AMT,
          (ISNULL(LS.Gross_Salary, 0) - ISNULL(ED.TOTAL_DED, 0)) AS NET_PAY
        FROM AuthorizedEmployees AE
        LEFT JOIN LatestSalary LS ON LS.Emp_Code = AE.EmpCode AND LS.rn = 1
        LEFT JOIN EmpDeductions ED ON ED.Emp_Id = AE.EmpCode
        LEFT JOIN MISC_MST MM_DEPT WITH (NOLOCK) ON MM_DEPT.Misc_Type = 68 AND MM_DEPT.Misc_Code = AE.DEPT_CODE AND ISNULL(MM_DEPT.Export_Type, 0) < 3
        LEFT JOIN godown_mst GM WITH (NOLOCK) ON GM.Godw_Code = ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM) AND ISNULL(GM.export_type, 0) < 3
        LEFT JOIN MISC_MST MM_LOC WITH (NOLOCK) ON MM_LOC.Misc_Type = 85 AND MM_LOC.Misc_Code = CAST(ISNULL(LS.Loc_Code, AE.LOC_CODE_NUM) AS VARCHAR(20)) AND ISNULL(MM_LOC.Export_Type, 0) < 3
      )
      SELECT
        SUM(GROSS_SALARY) AS OVERALL_GROSS,
        SUM(DEDUCTION_AMT) AS OVERALL_DED,
        SUM(NET_PAY) AS OVERALL_NET,
        COUNT(1) AS TOTAL_EMPLOYEES
      FROM EmployeeCompensation;
    `;

    const [[overallSalary]] = await Promise.all([
      sequelize.query(salaryQuery, { replacements })
    ]);

    const gross = Number(overallSalary[0]?.OVERALL_GROSS || 0);
    const deduction = Number(overallSalary[0]?.OVERALL_DED || 0);
    const netPay = Number(overallSalary[0]?.OVERALL_NET || (gross - deduction));

    return res.status(200).json({
      success: true,
      salary: {
        gross,
        deduction,
        netPay
      }
    });
  } catch (err) {
    console.error("❌ getManagerSalaryStats Error:", err);
    return res.status(500).json({ success: false, error: err.message });
  } finally {
    if (sequelize) await sequelize.close();
  }
};

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * 4. MODULAR ENDPOINT: ACTION ITEMS (Waiting On Me)
 * ══════════════════════════════════════════════════════════════════════════════
 */
exports.getManagerActionItems = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  try {
    const managerEmpCode = req.body.empcode || req.body.manager_empcode || req.body.EMPCODE || req.headers.empcode || "";
    const { Loc_code, dept_code } = req.body;
    let date = req.body.selectedDate || req.body.date || new Date().toISOString().split("T")[0];

    const { authorizedEmpCTE, replacements } = buildManagerScopeQuery({
      managerEmpCode,
      Loc_code,
      dept_code,
      selectedDate: date
    });

    const actionsQuery = `
      ${authorizedEmpCTE}
      SELECT
        AT.UTD AS REFERENCE_ID,
        AE.EmpCode AS EMP_CODE,
        AE.FULL_NAME AS EMP_NAME,
        AE.DESIGNATION,
        AE.DEPT_CODE,
        ISNULL(DEPT.Misc_Name, 'Department') AS DEPARTMENT,
        CAST(AT.DATEOFFICE AS DATE) AS ACTION_DATE,
        CASE
          WHEN AT.Leave_applied_on IS NOT NULL THEN 'LEAVE_APPROVAL'
          WHEN AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL THEN 'MISPUNCH_APPROVAL'
          ELSE 'GENERAL_APPROVAL'
        END AS ACTION_TYPE,
        CASE
          WHEN AT.Leave_applied_on IS NOT NULL THEN CONCAT('Leave Request: ', ISNULL(LM.Misc_Name, 'Leave'))
          WHEN AT.Mispunch_applied_on IS NOT NULL AND AT.Mispunch_out_applied_on IS NOT NULL THEN 'Mispunch In & Out Request'
          WHEN AT.Mispunch_applied_on IS NOT NULL THEN 'Mispunch In Request'
          WHEN AT.Mispunch_out_applied_on IS NOT NULL THEN 'Mispunch Out Request'
          ELSE 'Pending Approval'
        END AS TITLE,
        ISNULL(AT.SPL_REMARK, AT.reason) AS REMARKS,
        ISNULL(AT.Leave_applied_on, ISNULL(AT.Mispunch_applied_on, AT.Mispunch_out_applied_on)) AS APPLIED_ON,
        CASE
          WHEN AT.Appr_1_Stat IS NULL THEN 'Level 1 Pending'
          WHEN AT.Appr_2_Stat IS NULL THEN 'Level 2 Pending'
          WHEN AT.Appr_3_Stat IS NULL THEN 'Level 3 Pending'
          ELSE 'Pending'
        END AS PENDING_LEVEL
      FROM attendancetable AT WITH (NOLOCK)
      INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK) ON DEPT.Misc_Type = 68 AND DEPT.Misc_Code = AE.DEPT_CODE AND ISNULL(DEPT.Export_Type, 0) < 3
      LEFT JOIN MISC_MST LM WITH (NOLOCK) ON LM.Misc_Type = 92 AND LM.Misc_Code = AT.mipunch_reason AND ISNULL(LM.Export_Type, 0) < 3
      WHERE (
        (AT.Leave_applied_on IS NOT NULL AND (
          (AT.Appr_1_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_1_Code = :managerEmpCode))
          OR (AT.Appr_1_Stat = 1 AND AT.Appr_2_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_2_Code = :managerEmpCode))
          OR (AT.Appr_2_Stat = 1 AND AT.Appr_3_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_3_Code = :managerEmpCode))
        ))
        OR ((AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL) AND (
          (AT.Appr_1_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_1_Code = :managerEmpCode))
          OR (AT.Appr_1_Stat = 1 AND AT.Appr_2_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_2_Code = :managerEmpCode))
          OR (AT.Appr_2_Stat = 1 AND AT.Appr_3_Stat IS NULL AND (:managerEmpCode = '' OR AT.Appr_3_Code = :managerEmpCode))
        ))
      )
      AND ISNULL(AT.MAN_APPR, 'N') <> 'Y'
      AND ISNULL(AT.MAN_REJ, 'N') <> 'Y'
      ORDER BY APPLIED_ON DESC;
    `;

    const [[actionItems]] = await Promise.all([
      sequelize.query(actionsQuery, { replacements })
    ]);

    return res.status(200).json({ success: true, count: actionItems.length, actions: actionItems });
  } catch (err) {
    console.error("❌ getManagerActionItems Error:", err);
    return res.status(500).json({ success: false, error: err.message });
  } finally {
    if (sequelize) await sequelize.close();
  }
};

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * 5. MODULAR ENDPOINT: WATCHLIST & EXCEPTIONS
 * ══════════════════════════════════════════════════════════════════════════════
 */
exports.getManagerWatchlist = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  try {
    const managerEmpCode = req.body.empcode || req.body.manager_empcode || req.body.EMPCODE || req.headers.empcode || "";
    const { Loc_code, dept_code } = req.body;
    let date = req.body.selectedDate || req.body.date || new Date().toISOString().split("T")[0];
    const selectedDateObj = new Date(date);
    const selectedMonth = req.body.month || (selectedDateObj.getMonth() + 1);
    const selectedYear = req.body.year || selectedDateObj.getFullYear();

    const { authorizedEmpCTE, replacements } = buildManagerScopeQuery({
      managerEmpCode,
      Loc_code,
      dept_code,
      selectedDate: date
    });
    replacements.selectedMonth = selectedMonth;
    replacements.selectedYear = selectedYear;

    const shiftStartExpr = `
      DATEADD(
        MINUTE,
        (FLOOR(AT.SHIFTSTARTTIME) * 60) + ROUND((AT.SHIFTSTARTTIME - FLOOR(AT.SHIFTSTARTTIME)) * 60, 0),
        CAST(AT.DATEOFFICE AS DATETIME)
      )
    `;

    const watchlistQuery = `
      ${authorizedEmpCTE},
      MonthlyMispunchSummary AS (
        SELECT
          AT.Emp_Code,
          COUNT(1) AS MISPUNCH_COUNT
        FROM attendancetable AT WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
        WHERE (MONTH(AT.DATEOFFICE) = :selectedMonth AND YEAR(AT.DATEOFFICE) = :selectedYear)
          AND (AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL)
        GROUP BY AT.Emp_Code
        HAVING COUNT(1) >= 2
      ),
      LateSummary AS (
        SELECT
          AT.Emp_Code,
          COUNT(1) AS LATE_COUNT
        FROM attendancetable AT WITH (NOLOCK)
        INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
        WHERE CAST(AT.DATEOFFICE AS DATE) BETWEEN DATEADD(DAY, -7, :selectedDate) AND :selectedDate
          AND AT.IN1 IS NOT NULL
          AND AT.IN1 > ${shiftStartExpr}
        GROUP BY AT.Emp_Code
        HAVING COUNT(1) >= 2
      )
      SELECT
        AE.EmpCode AS employeeId,
        AE.FULL_NAME AS employeeName,
        ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS department,
        ISNULL(GM.Godw_Name, AE.LOC_CODE_STR) AS location,
        'HIGH_MISPUNCH' AS issueType,
        CONCAT(MS.MISPUNCH_COUNT, ' mispunches recorded in current month') AS message,
        CAST(MS.MISPUNCH_COUNT AS VARCHAR(20)) AS value,
        'MEDIUM' AS severity,
        AE.EmpCode AS referenceId
      FROM MonthlyMispunchSummary MS
      INNER JOIN AuthorizedEmployees AE ON MS.Emp_Code = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK) ON DEPT.Misc_Type = 68 AND DEPT.Misc_Code = AE.DEPT_CODE
      LEFT JOIN godown_mst GM WITH (NOLOCK) ON GM.Godw_Code = AE.LOC_CODE_NUM

      UNION ALL

      SELECT
        AE.EmpCode AS employeeId,
        AE.FULL_NAME AS employeeName,
        ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS department,
        ISNULL(GM.Godw_Name, AE.LOC_CODE_STR) AS location,
        'FREQUENT_LATE' AS issueType,
        CONCAT(LS.LATE_COUNT, ' late arrivals in last 7 days') AS message,
        CAST(LS.LATE_COUNT AS VARCHAR(20)) AS value,
        'HIGH' AS severity,
        AE.EmpCode AS referenceId
      FROM LateSummary LS
      INNER JOIN AuthorizedEmployees AE ON LS.Emp_Code = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK) ON DEPT.Misc_Type = 68 AND DEPT.Misc_Code = AE.DEPT_CODE
      LEFT JOIN godown_mst GM WITH (NOLOCK) ON GM.Godw_Code = AE.LOC_CODE_NUM;
    `;

    const [[watchlist]] = await Promise.all([
      sequelize.query(watchlistQuery, { replacements })
    ]);

    return res.status(200).json({ success: true, count: watchlist.length, watchlist });
  } catch (err) {
    console.error("❌ getManagerWatchlist Error:", err);
    return res.status(500).json({ success: false, error: err.message });
  } finally {
    if (sequelize) await sequelize.close();
  }
};

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * 6. CEO DASHBOARD API (Preserved for Reference & Compatibility)
 * ══════════════════════════════════════════════════════════════════════════════
 */
exports.getDashboardData = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  try {
    const { Loc_code, selectedDate, shiftstarttime, dept_code } = req.body;

    let date = selectedDate;
    if (!date) {
      const today = new Date();
      date = today.toISOString().split("T")[0];
    }

    let locFilter = "";
    const replacements = { selectedDate: date };

    if (Loc_code && Loc_code !== "") {
      if (Array.isArray(Loc_code) && Loc_code.length > 0) {
        const locParams = Loc_code.map((_, idx) => `:loc_${idx}`);
        Loc_code.forEach((val, idx) => {
          replacements[`loc_${idx}`] = String(val);
        });
        locFilter = `AND (EM.LOCATION IN (${locParams.join(",")}) OR CAST(EM.Loc_Code AS VARCHAR(20)) IN (${locParams.join(",")}))`;
      } else {
        replacements.locSingle = String(Loc_code);
        locFilter = `AND (EM.LOCATION = :locSingle OR CAST(EM.Loc_Code AS VARCHAR(20)) = :locSingle)`;
      }
    }

    let shiftFilter = "";
    if (shiftstarttime !== undefined && shiftstarttime !== null && shiftstarttime !== "") {
      replacements.shiftStartTime = String(shiftstarttime);
      shiftFilter = `AND AT.SHIFTSTARTTIME = :shiftStartTime`;
    }

    const shiftStartExpr = `
      DATEADD(
        MINUTE,
        (FLOOR(AT.SHIFTSTARTTIME) * 60) + ROUND((AT.SHIFTSTARTTIME - FLOOR(AT.SHIFTSTARTTIME)) * 60, 0),
        CAST(AT.DATEOFFICE AS DATETIME)
      )
    `;

    let deptFilter = "";
    if (dept_code && dept_code !== "") {
      replacements.deptCode = String(dept_code);
      deptFilter = `AND EM.Division = :deptCode`;
    }

    const shiftQuery = `
      SELECT DISTINCT
        AT.SHIFTSTARTTIME,
        AT.SHIFTENDTIME,
        AT.SHIFT
      FROM attendancetable AT WITH (NOLOCK)
      INNER JOIN EMPLOYEEMASTER EM WITH (NOLOCK)
        ON AT.Emp_Code = EM.EmpCode
       AND EM.LASTWOR_DATE IS NULL
      WHERE CAST(AT.DATEOFFICE AS DATE) = :selectedDate
        ${locFilter}
      ORDER BY AT.SHIFTSTARTTIME
    `;

    const graphQuery = `
      ;WITH DateRange AS (
          SELECT CAST(:selectedDate AS DATE) AS dt
          UNION ALL
          SELECT DATEADD(DAY, -1, dt) FROM DateRange WHERE dt > DATEADD(DAY, -7, :selectedDate)
      ),
      AT_DEDUP AS (
          SELECT
              AT.*,
              ROW_NUMBER() OVER (
                  PARTITION BY AT.Emp_Code, CAST(AT.DATEOFFICE AS DATE)
                  ORDER BY CASE WHEN AT.IN1 IS NOT NULL THEN 0 ELSE 1 END, AT.IN1 DESC
              ) AS rn
          FROM attendancetable AT WITH (NOLOCK)
          WHERE CAST(AT.DATEOFFICE AS DATE) BETWEEN DATEADD(DAY, -7, :selectedDate) AND :selectedDate
          ${shiftFilter}
      ),
      STATUS_CALC AS (
          SELECT
              DR.dt AS DATE_LABEL,
              CASE
                  WHEN AT.Emp_Code IS NULL THEN 'ABSENT'
                  WHEN DR.dt = CAST(GETDATE() AS DATE) THEN
                      CASE WHEN AT.FLAG = 'P' THEN 'PRESENT' ELSE 'ABSENT' END
                  ELSE
                      CASE WHEN AT.STATUS = 'A' THEN 'ABSENT' ELSE 'PRESENT' END
              END AS PUNCH_STATUS
          FROM DateRange DR
          CROSS JOIN EMPLOYEEMASTER EM
          LEFT JOIN AT_DEDUP AT
              ON AT.Emp_Code = EM.EmpCode
             AND CAST(AT.DATEOFFICE AS DATE) = DR.dt
             AND AT.rn = 1
          WHERE
              EM.LASTWOR_DATE IS NULL
              ${locFilter}
              ${deptFilter}
      )
      SELECT
          DATE_LABEL,
          DATENAME(WEEKDAY, DATE_LABEL) AS DAY_NAME,
          COUNT(CASE WHEN PUNCH_STATUS = 'PRESENT' THEN 1 END) AS PRESENT,
          COUNT(CASE WHEN PUNCH_STATUS = 'ABSENT' THEN 1 END) AS ABSENT
      FROM STATUS_CALC
      GROUP BY DATE_LABEL
      ORDER BY DATE_LABEL ASC
      OPTION (MAXRECURSION 10)
    `;

    const deptStatsQuery = `
      SELECT
        MM.Misc_Code AS DEPT_CODE,
        MM.Misc_Name AS DEPT_NAME,
        COUNT(EM.EmpCode) AS TOTAL,
        COUNT(
          CASE
            WHEN CAST(AT.DATEOFFICE AS DATE) = :selectedDate AND AT.FLAG = 'P' THEN 1
            WHEN CAST(AT.DATEOFFICE AS DATE) <> :selectedDate AND AT.IN1 IS NOT NULL THEN 1
          END
        ) AS PRESENT,
        COUNT(
          CASE
            WHEN CAST(AT.DATEOFFICE AS DATE) = :selectedDate AND (AT.FLAG = 'A' OR AT.FLAG IS NULL) THEN 1
            WHEN CAST(AT.DATEOFFICE AS DATE) <> :selectedDate AND (AT.IN1 IS NULL OR AT.STATUS = 'A') THEN 1
          END
        ) AS ABSENT
      FROM MISC_MST MM WITH (NOLOCK)
      LEFT JOIN EMPLOYEEMASTER EM WITH (NOLOCK)
        ON EM.Division = MM.Misc_Code
       AND EM.LASTWOR_DATE IS NULL
      LEFT JOIN attendancetable AT WITH (NOLOCK)
        ON AT.Emp_Code = EM.EmpCode
       AND CAST(AT.DATEOFFICE AS DATE) = :selectedDate
       ${shiftFilter}
      WHERE
        MM.Misc_Type = 68
        AND ISNULL(MM.Export_Type,0) < 3
        ${locFilter}
      GROUP BY MM.Misc_Code, MM.Misc_Name
      ORDER BY MM.Misc_Name
    `;

    const employeeQuery = `
      SELECT
        EM.EmpCode AS EMP_CODE,
        EM.EmpFirstName AS EMP_NAME,
        EM.EMPLOYEEDesignation AS DESIGNATION,
        EM.EmployeeType AS EMPLOYEE_TYPE,
        DEPT.Misc_Name AS DEPARTMENT,
        CONVERT(VARCHAR(8), AT.IN1, 108) AS LOGIN_TIME,
        CONVERT(VARCHAR(8), AT.OUT1, 108) AS LOGOUT_TIME,
        AT.STATUS AS ATTENDANCE_STATUS,
        CASE
          WHEN AT.Emp_Code IS NULL THEN 'ABSENT'
          WHEN CAST(GETDATE() AS DATE) = :selectedDate THEN
            CASE WHEN AT.FLAG = 'P' THEN 'PRESENT' ELSE 'ABSENT' END
          ELSE
            CASE WHEN AT.STATUS = 'A' THEN 'ABSENT' ELSE 'PRESENT' END
        END AS PUNCH_STATUS,
        ED.DOC_PATH AS PHOTO_PATH
      FROM EMPLOYEEMASTER EM WITH (NOLOCK)
      LEFT JOIN attendancetable AT WITH (NOLOCK)
        ON AT.Emp_Code = EM.EmpCode
       AND CAST(AT.DATEOFFICE AS DATE) = :selectedDate
       ${shiftFilter}
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK)
        ON DEPT.Misc_Type = 68
       AND DEPT.Misc_Code = EM.Division
       AND ISNULL(DEPT.Export_Type,0) < 3
      LEFT JOIN EMP_DOCS ED WITH (NOLOCK)
        ON ED.EMP_CODE = EM.EmpCode
       AND ED.columndoc_type = 'EMPLOYEE'
       AND ED.Seq_No = 1
       AND ISNULL(ED.export_type,0) < 3
      WHERE 1=1
        AND EM.LASTWOR_DATE IS NULL
        ${locFilter}
        ${deptFilter}
      ORDER BY EM.EmpFirstName
    `;

    const pendingLeaveListQuery = `
      SELECT
        EM.EmpCode AS EMP_CODE,
        CONCAT(EM.EmpFirstName, ' ', EM.EmpLastName) AS EMP_NAME,
        DEPT.Misc_Name AS DEPARTMENT,
        CAST(AT.DATEOFFICE AS DATE) AS LEAVE_DATE,
        AT.Leave_applied_on AS APPLIED_ON,
        LM.Misc_Name AS LEAVE_TYPE,
        'Pending' AS STATUS
      FROM attendancetable AT WITH (NOLOCK)
      INNER JOIN EMPLOYEEMASTER EM WITH (NOLOCK)
        ON EM.EmpCode = AT.Emp_Code
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK)
        ON DEPT.Misc_Type = 68
       AND DEPT.Misc_Code = EM.Division
      LEFT JOIN MISC_MST LM WITH (NOLOCK)
        ON LM.Misc_Type = 92
       AND LM.Assessable_Column IN (1,2,3)
       AND LM.Misc_Code = AT.mipunch_reason
      WHERE
        CAST(AT.DATEOFFICE AS DATE) = :selectedDate
        ${locFilter}
        ${shiftFilter}
        AND AT.Leave_applied_on IS NOT NULL
        AND (
          (AT.Appr_1_Stat IS NULL AND AT.Appr_2_Stat IS NULL AND AT.Appr_3_Stat IS NULL)
          OR (AT.Appr_1_Stat = 1 AND (AT.Appr_2_Code IS NOT NULL AND AT.Appr_2_Stat IS NULL))
          OR (AT.Appr_2_Stat = 1 AND (AT.Appr_3_Code IS NOT NULL AND AT.Appr_3_Stat IS NULL))
        )
        AND ISNULL(AT.MAN_APPR,'N') <> 'Y'
        AND ISNULL(AT.MAN_REJ,'N') <> 'Y'
      ORDER BY EM.EmpFirstName
    `;

    const [
      [shifts],
      [graphData],
      [deptStats],
      [employees],
      [pendingLeaves]
    ] = await Promise.all([
      sequelize.query(shiftQuery, { replacements }),
      sequelize.query(graphQuery, { replacements }),
      sequelize.query(deptStatsQuery, { replacements }),
      sequelize.query(employeeQuery, { replacements }),
      sequelize.query(pendingLeaveListQuery, { replacements })
    ]);

    const BASE_URL = process.env.DOC_BASE_URL || "";
    const employeeList = employees.map(emp => ({
      ...emp,
      PHOTO_URL: emp.PHOTO_PATH ? `${BASE_URL}${emp.PHOTO_PATH}` : null,
    }));
    const employeesData = {
      all: employeeList,
      present: employeeList.filter(e => e.PUNCH_STATUS === "PRESENT"),
      absent: employeeList.filter(e => e.PUNCH_STATUS === "ABSENT"),
    };

    return res.json({
      success: true,
      date: date,
      result: {
        shifts,
        graphData,
        deptStats,
        employees: employeesData,
        pendingLeaves
      },
    });

  } catch (err) {
    console.error("❌ getDashboardData Error:", err);
    return res.status(500).json({ success: false, error: err.message });
  } finally {
    if (sequelize) await sequelize.close();
  }
};

/**
 * ══════════════════════════════════════════════════════════════════════════════
 * 7. "NEEDS YOU TODAY" WORK ITEMS ENDPOINT
 * ══════════════════════════════════════════════════════════════════════════════
 * Returns manager's daily actionable items:
 * 1. Mispunch Approvals (MAN_APPR = 'N' AND MAN_REJ = 'N')
 * 2. Shift Request Approvals (Appr_1_Stat IS NULL)
 */
exports.getNeedsYouToday = async (req, res) => {
  const sequelize = await dbname(req, req.headers.compcode);
  try {
    const managerEmpCode = req.body.empcode || req.body.manager_empcode || req.body.EMPCODE || req.headers.empcode || "";
    const { Loc_code, dept_code } = req.body;
    let date = req.body.selectedDate || req.body.date || new Date().toISOString().split("T")[0];
    const selectedDateObj = new Date(date);
    const selectedMonth = req.body.month || (selectedDateObj.getMonth() + 1);
    const selectedYear = req.body.year || selectedDateObj.getFullYear();

    const { authorizedEmpCTE, replacements } = buildManagerScopeQuery({
      managerEmpCode,
      Loc_code,
      dept_code,
      selectedDate: date
    });
    replacements.selectedMonth = selectedMonth;
    replacements.selectedYear = selectedYear;

    // ──────────────────────────────────────────────────────────────────────────
    // 1. MISPUNCH APPROVALS QUERY (Strictly Current Active Month)
    // ──────────────────────────────────────────────────────────────────────────
    const mispunchQuery = `
      ${authorizedEmpCTE}
      SELECT
        AT.UTD AS id,
        AE.EmpCode AS employeeCode,
        AE.FULL_NAME AS employeeName,
        AE.DESIGNATION AS designation,
        AE.EMPLOYEE_TYPE AS employeeType,
        ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS department,
        CAST(AT.DATEOFFICE AS DATE) AS attendanceDate,
        AT.Mispunch_applied_on AS appliedOnIn,
        AT.Mispunch_out_applied_on AS appliedOnOut,
        ISNULL(AT.Mispunch_applied_on, AT.Mispunch_out_applied_on) AS appliedOn,
        CASE
          WHEN AT.Mispunch_applied_on IS NOT NULL AND AT.Mispunch_out_applied_on IS NOT NULL THEN 'Mispunch In & Out'
          WHEN AT.Mispunch_applied_on IS NOT NULL THEN 'Mispunch In'
          WHEN AT.Mispunch_out_applied_on IS NOT NULL THEN 'Mispunch Out'
          ELSE 'Mispunch'
        END AS mispunchType,
        ISNULL(MM.Misc_Name, 'Mispunch Regularization') AS reason,
        ISNULL(AT.SPL_REMARK, AT.reason) AS remarks,
        CONVERT(VARCHAR(8), AT.IN1, 108) AS inTime,
        CONVERT(VARCHAR(8), AT.OUT1, 108) AS outTime,
        CONVERT(VARCHAR(8), AT.App_in1, 108) AS requestedInTime,
        CONVERT(VARCHAR(8), AT.App_out1, 108) AS requestedOutTime,
        'Pending' AS status,
        DATEDIFF(DAY, CAST(ISNULL(AT.Mispunch_applied_on, AT.DATEOFFICE) AS DATE), CAST(:selectedDate AS DATE)) AS daysPending,
        ED.DOC_PATH AS photoPath
      FROM attendancetable AT WITH (NOLOCK)
      INNER JOIN AuthorizedEmployees AE ON AT.Emp_Code = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK)
        ON DEPT.Misc_Type = 68
       AND DEPT.Misc_Code = AE.DEPT_CODE
       AND ISNULL(DEPT.Export_Type, 0) < 3
      LEFT JOIN MISC_MST MM WITH (NOLOCK)
        ON MM.Misc_Type = 92
       AND MM.Misc_Code = AT.mipunch_reason
       AND ISNULL(MM.Export_Type, 0) < 3
      LEFT JOIN EMP_DOCS ED WITH (NOLOCK)
        ON ED.EMP_CODE = AE.EmpCode
       AND ED.columndoc_type = 'EMPLOYEE'
       AND ED.Seq_No = 1
       AND ISNULL(ED.export_type, 0) < 3
      WHERE (AT.Mispunch_applied_on IS NOT NULL OR AT.Mispunch_out_applied_on IS NOT NULL OR AT.mipunch_reason IS NOT NULL)
        -- Strictly Current Active Month Filter (No previous month data)
        AND MONTH(AT.DATEOFFICE) = :selectedMonth
        AND YEAR(AT.DATEOFFICE) = :selectedYear
        -- Pending Condition
        AND ISNULL(AT.MAN_APPR, 'N') = 'N'
        AND ISNULL(AT.MAN_REJ, 'N') = 'N'
        -- Manager Approver filter
        AND (
          :managerEmpCode = ''
          OR AT.Appr_1_Code = :managerEmpCode
          OR (AT.Appr_1_Stat = 1 AND AT.Appr_2_Code = :managerEmpCode AND AT.Appr_2_Stat IS NULL)
          OR (AT.Appr_1_Code IS NULL)
        )
      ORDER BY ISNULL(AT.Mispunch_applied_on, AT.DATEOFFICE) DESC;
    `;

    // ──────────────────────────────────────────────────────────────────────────
    // 2. SHIFT REQUEST APPROVALS QUERY (Strictly Current Active Month)
    // ──────────────────────────────────────────────────────────────────────────
    const shiftRequestQuery = `
      ${authorizedEmpCTE}
      SELECT
        SR.Tran_id AS id,
        AE.EmpCode AS employeeCode,
        AE.FULL_NAME AS employeeName,
        AE.DESIGNATION AS designation,
        AE.EMPLOYEE_TYPE AS employeeType,
        ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS department,
        SR.ATN_DATE_FROM AS dateFrom,
        SR.ATN_DATE_TO AS dateTo,
        SR.Current_SHIFT AS currentShift,
        SR.Req_SHIFT AS requestedShift,
        ISNULL(SR.Appr_1_Rem, '') AS remarks,
        'Pending' AS status,
        ED.DOC_PATH AS photoPath
      FROM Shift_Request SR WITH (NOLOCK)
      INNER JOIN AuthorizedEmployees AE ON SR.EMPCODE = AE.EmpCode
      LEFT JOIN MISC_MST DEPT WITH (NOLOCK)
        ON DEPT.Misc_Type = 68
       AND DEPT.Misc_Code = AE.DEPT_CODE
       AND ISNULL(DEPT.Export_Type, 0) < 3
      LEFT JOIN EMP_DOCS ED WITH (NOLOCK)
        ON ED.EMP_CODE = AE.EmpCode
       AND ED.columndoc_type = 'EMPLOYEE'
       AND ED.Seq_No = 1
       AND ISNULL(ED.export_type, 0) < 3
      WHERE SR.Appr_1_Stat IS NULL
        -- Strictly Current Active Month Filter (No previous month data)
        AND (
          (MONTH(SR.ATN_DATE_FROM) = :selectedMonth AND YEAR(SR.ATN_DATE_FROM) = :selectedYear)
          OR (MONTH(SR.ATN_DATE_TO) = :selectedMonth AND YEAR(SR.ATN_DATE_TO) = :selectedYear)
        )
        AND (:managerEmpCode = '' OR SR.Appr_1_Code = :managerEmpCode)
        AND ISNULL(SR.Fin_Appr, 0) = 0
      ORDER BY SR.Tran_id DESC;
    `;

    const [mispunchRows] = await sequelize.query(mispunchQuery, { replacements });

    let shiftRequestRows = [];
    try {
      const [shifts] = await sequelize.query(shiftRequestQuery, { replacements });
      shiftRequestRows = shifts;
    } catch (shiftErr) {
      try {
        const altShiftQuery = `
          ${authorizedEmpCTE}
          SELECT
            WO.Tran_id AS id,
            AE.EmpCode AS employeeCode,
            AE.FULL_NAME AS employeeName,
            AE.DESIGNATION AS designation,
            ISNULL(DEPT.Misc_Name, AE.DEPT_CODE) AS department,
            WO.Change_Date AS dateFrom,
            WO.Change_Date AS dateTo,
            WO.Prev_WO AS currentShift,
            WO.New_WO AS requestedShift,
            ISNULL(WO.Remark, '') AS remarks,
            'Pending' AS status
          FROM WO_Changes_dtl WO WITH (NOLOCK)
          INNER JOIN AuthorizedEmployees AE ON WO.EMPCODE = AE.EmpCode
          LEFT JOIN MISC_MST DEPT WITH (NOLOCK) ON DEPT.Misc_Type = 68 AND DEPT.Misc_Code = AE.DEPT_CODE AND ISNULL(DEPT.Export_Type, 0) < 3
          WHERE WO.Appr_1_Stat IS NULL
            -- Strictly Current Active Month Filter (No previous month data)
            AND MONTH(WO.Change_Date) = :selectedMonth 
            AND YEAR(WO.Change_Date) = :selectedYear
            AND (:managerEmpCode = '' OR WO.Appr_1_Code = :managerEmpCode)
            AND ISNULL(WO.Fin_Appr, 0) = 0
          ORDER BY WO.Tran_id DESC;
        `;
        const [altShifts] = await sequelize.query(altShiftQuery, { replacements });
        shiftRequestRows = altShifts;
      } catch (e) {
        shiftRequestRows = [];
      }
    }

    const BASE_URL = process.env.DOC_BASE_URL || "";

    const mispunchItems = mispunchRows.map(item => ({
      ...item,
      photoUrl: item.photoPath ? `${BASE_URL}${item.photoPath}` : null
    }));

    const shiftItems = shiftRequestRows.map(item => ({
      ...item,
      photoUrl: item.photoPath ? `${BASE_URL}${item.photoPath}` : null
    }));

    const maxMispunchDays = mispunchItems.reduce((max, i) => Math.max(max, i.daysPending || 0), 0);
    const mispunchSubtitle = mispunchItems.length > 0
      ? (maxMispunchDays > 0 ? `Oldest is ${maxMispunchDays} day${maxMispunchDays > 1 ? "s" : ""} old` : "Pending for today")
      : "No pending mispunches";

    const shiftSubtitle = shiftItems.length > 0
      ? `${shiftItems.length} shift change request${shiftItems.length > 1 ? "s" : ""} pending`
      : "No pending shift requests";

    return res.status(200).json({
      success: true,
      date: date,
      data: {
        totalPendingCount: mispunchItems.length + shiftItems.length,
        mispunch: {
          title: "Mispunch approvals",
          count: mispunchItems.length,
          subtitle: mispunchSubtitle,
          items: mispunchItems
        },
        shiftRequest: {
          title: "Shift change requests",
          count: shiftItems.length,
          subtitle: shiftSubtitle,
          items: shiftItems
        }
      }
    });

  } catch (err) {
    console.error("❌ getNeedsYouToday Error:", err);
    return res.status(500).json({ success: false, error: err.message });
  } finally {
    if (sequelize) await sequelize.close();
  }
};


exports.gethrentry = async function (req, res) {
  let sequelize;
  try {
    sequelize = await dbname(req, req.headers.compcode);

    const branch = await sequelize.query(
      `SELECT *,(select top 1 misc_name from misc_mst where misc_Code = NEW_JOINING.loc_code and misc_type = 85) as locationname FROM NEW_JOINING order by tran_id desc;`,
    );

    const cardData = await sequelize.query(
      `SELECT 
        DESIGNATION,
        (SELECT TOP 1 misc_name FROM misc_mst WHERE misc_Code = MIN(NEW_JOINING.loc_code) AND misc_type = 85) AS locationname,
        MIN(NEW_JOINING.loc_code) AS loc_code,
        COUNT(*) AS total,
        SUM(CASE WHEN ADDRESS IS NOT NULL AND LTRIM(RTRIM(ADDRESS)) <> '' THEN 1 ELSE 0 END) AS filled,
        SUM(CASE WHEN ADDRESS IS NULL OR LTRIM(RTRIM(ADDRESS)) = '' THEN 1 ELSE 0 END) AS pending,
        CASE 
          WHEN COUNT(*) > 0 AND SUM(CASE WHEN ADDRESS IS NOT NULL AND LTRIM(RTRIM(ADDRESS)) <> '' THEN 1 ELSE 0 END) = COUNT(*) THEN 'Filled'
          ELSE 'Open'
        END AS status
      FROM NEW_JOINING 
      WHERE DESIGNATION IS NOT NULL AND LTRIM(RTRIM(DESIGNATION)) <> ''
      GROUP BY DESIGNATION
      ORDER BY total DESC;`,
    );

    res.status(200).send({
      data: branch[0],
      cardData: cardData[0]
    });
  } catch (e) {
    console.log(e);
    res.status(500).send({ message: e.message });
  } finally {
    if (sequelize) {
      await sequelize.close();
    }
  }
};

exports.interviewcanidates = async function (req, res) {
  let sequelize;
  try {
    sequelize = await dbname(req, req.headers.compcode);
    const { loc_code } = req.body;

    // Location filter safely handle karein
    let locCondition = "";
    if (loc_code !== undefined && loc_code !== null && String(loc_code).trim() !== "") {
      const formattedLocs = String(loc_code)
        .split(",")
        .map((l) => `'${l.trim()}'`)
        .join(",");
      locCondition = `AND LOC_CODE IN (${formattedLocs})`;
    }

    const branch = await sequelize.query(`
SELECT
    nj.CITY as CITY1,
    nj.REJECTED_BY,
    CAST(nj.REJECTION_DATE AS DATE) as REJECTION_DATE,
    (select top 1 Misc_Name from Misc_Mst where Misc_Code = nj.STATE and Misc_Type = 3) AS STATE1,
    (select top 1 Misc_Name from Misc_Mst where Misc_Code = nj.RELIGION and Misc_Type = 603) AS RELIGION1,
    (select top 1 Misc_Name from Misc_Mst where Misc_Code = nj.LOC_CODE and Misc_Type = 85) AS LOC_CODE1,
    CONVERT(varchar, nj.APPLICATION_DATE, 105) as APPLICATION_DATE1,
    CONVERT(varchar, nj.DOB, 105) as DOB1,
    CONVERT(varchar, nj.DOM, 105) as DOM1,
    (select top 1 CONCAT(em1.empfirstname, ' ', em1.emplastname) from EMPLOYEEMASTER em1 where em1.EMPCODE=nj.INTR1BY) as employeename1,
    (select top 1 CONCAT(em2.empfirstname, ' ', em2.emplastname) from EMPLOYEEMASTER em2 where em2.EMPCODE=nj.INTR2BY) as employeename2,
    (select top 1 CONCAT(em3.empfirstname, ' ', em3.emplastname) from EMPLOYEEMASTER em3 where em3.EMPCODE=nj.INTR3BY) as employeename3,
    (select top 1 CONCAT(em4.empfirstname, ' ', em4.emplastname) from EMPLOYEEMASTER em4 where em4.EMPCODE=nj.INTR4BY) as employeename4,
    nj.*
FROM
    NEW_JOINING nj
WHERE
    1=1
    ${locCondition}
ORDER BY 
    nj.TRAN_ID DESC
`);

    const SRNOS = branch[0].map((abcd) => abcd.TRAN_ID);
    if (!SRNOS?.length) {
      return res.status(200).send([]);
    }
    const interviewSideData = await sequelize.query(
      `SELECT * FROM Interview_SideTables where SRNO in (${SRNOS.join(",")})`,
    );
    const ImageData = await sequelize.query(
      `select * from DOC_UPLOAD where Doc_Type='NCR' and Export_type < 3 and TRAN_ID in (${SRNOS.join(
        ",",
      )})`,
    );

    const fieldsToKeepByType = {
      1: [
        // EmpExperience
        "Emp_Company",
        "Emp_Designation",
        "Emp_Responsibility",
        "Emp_From_Date",
        "Emp_To_Date",
        "Emp_Settlement_Done",
        "Emp_Drawn_Salary",
        "Emp_Leaving_Reason",
      ],
      2: [
        // EmpEdu
        "Emp_Degree",
        "Emp_Board",
        "Emp_College",
        "Emp_Passing_year",
        "Emp_Percentage",
      ],
      3: [
        // EmpItSkill
        "Emp_Tool",
        "Emp_Version",
        "Emp_Proficiency",
        "Emp_Last_Used",
        "Emp_Experience",
      ],
      4: [
        // EmpLang
        "Emp_Language",
        "Emp_Language_Understand",
        "Emp_Language_Speak",
        "Emp_Language_Read",
        "Emp_Language_Write",
      ],
      5: [
        // References
        "Emp_Ref_Name",
        "Emp_Ref_Occup",
        "Emp_Ref_Address",
        "Emp_Ref_Mobile",
        "Emp_Ref_emailid",
        "Emp_Ref_relation",
      ],
      6: [
        // EmpNominee
        "Nominee_Name",
        "Member_Name",
        "Relation",
        "Percentage",
        "Is_Minor",
      ],
    };

    const Tbl_Type = {
      EmpExperience: 1,
      EmpEdu: 2,
      EmpItSkill: 3,
      EmpLang: 4,
      References: 5,
      EmpNominee: 6,
    };

    const processDataForAllSRNOs = (interviewData, ImageData) => {
      // Grouping the data by SRNO
      const groupedData = interviewData.reduce((acc, row) => {
        const { SRNO, Tbl_Type, ...rest } = row;
        const fieldsToKeep = fieldsToKeepByType[Tbl_Type] || [];

        // Filter out null values and keep only relevant fields
        const filteredData = fieldsToKeep.reduce((obj, field) => {
          if (rest[field] !== null && rest[field] !== undefined) {
            obj[field] = rest[field];
          }
          return obj;
        }, {});

        // Only push if there is any relevant data
        if (Object.keys(filteredData).length > 0) {
          if (!acc[SRNO]) {
            acc[SRNO] = {
              EmpEdu: [],
              EmpLang: [],
              EmpItSkill: [],
              EmpExperience: [],
              References: [],
              EmpFamily: [],
              EmpNominee: [],
            };
          }

          // Mapping data to the correct array based on Tbl_Type
          switch (Tbl_Type) {
            case 1:
              acc[SRNO].EmpExperience.push(filteredData);
              break;
            case 2:
              acc[SRNO].EmpEdu.push(filteredData);
              break;
            case 3:
              acc[SRNO].EmpItSkill.push(filteredData);
              break;
            case 4:
              acc[SRNO].EmpLang.push(filteredData);
              break;
            case 5:
              acc[SRNO].References.push(filteredData);
              break;
            case 6:
              acc[SRNO].EmpNominee.push(filteredData);
              break;
          }
        }

        return acc;
      }, {});

      const groupedImageData = ImageData.reduce((acc, row) => {
        const { TRAN_ID, ...imageFields } = row;

        // Ensure there is an array to store images for each SRNO
        if (!acc[TRAN_ID]) {
          acc[TRAN_ID] = [];
        }

        // Push the image data into the array
        acc[TRAN_ID].push(imageFields);

        return acc;
      }, {});

      // Convert groupedData into an array if needed
      const finalData = branch[0].map((srno) => ({
        ...srno,
        ...groupedData[srno.TRAN_ID],
        IMAGES: groupedImageData[srno.TRAN_ID],
      }));

      return finalData;
    };

    // Call the function and get the processed data
    const result = processDataForAllSRNOs(interviewSideData[0], ImageData[0]);

    res.status(200).send(result);
  } catch (e) {
    console.log(e);
    return res.status(500).send({ message: "Error" });
  } finally {
    if (sequelize) {
      await sequelize.close();
    }
  }
};

exports.shortlistcandidate = async function (req, res) {
  let sequelize;
  let t;
  try {
    if (!req.body.tran_id) {
      return res.status(400).send({
        status: false,
        message: "tran_id is mandatory",
      });
    }
    sequelize = await dbname(req, req.headers.compcode);
    t = await sequelize.transaction();

    const randomUUID = uuidv4();

    await sequelize.query(
      `
      UPDATE NEW_JOINING
      SET int_status = 2, unique_id = '${randomUUID}'
      WHERE tran_id = ${req.body.tran_id}
    `,
      { transaction: t },
    );

    await sequelize.query(
      `
      INSERT INTO SHORTLISTED_CANDIDATE (
        SKILLS, SRNO, EMPFIRSTNAME, MOBILE_NO, MOBILENO, CURRENTADDRESS1,
        PPINCODE, pState, PCITY, BASICQUALIFICATION, FATHERNAME,
        GENDER, MOTHERNAME, LOCATION, EMPLOYEEDESIGNATION, ALTERNET_MAIL,
        UID_NO, DOB, DOM, Export_Type, EMPCODE, SERVERid ,
         CREATED_BY, CREATED_ON , CLUSTER , CHANNEL,SUB_SOURCE,SOURCE_OF_REG
      )
      SELECT
        nj.SKILLS, nj.tran_id, nj.NAME, nj.MOB_NO, nj.WHATSAPP_NO, nj.ADDRESS,
        nj.PINCODE, nj.STATE, '', nj.HIGH_QUAL, nj.FATHERS_NAME,
        nj.GENDER, nj.MOTHERS_NAME, nj.LOC_CODE, nj.DESIGNATION,
        nj.EMAIL, nj.AADHAR_NO, nj.DOB, nj.DOM, nj.INT_status, '', '' ,
        '${req.headers?.name}' , getDate() , nj.CLUSTER,nj.CHANNEL,nj.SUB_SOURCE,nj.SOURCE_OF_REG
      FROM NEW_JOINING nj
      WHERE nj.tran_id = ${req.body.tran_id}
      AND NOT EXISTS (
        SELECT 1
        FROM SHORTLISTED_CANDIDATE sc
        WHERE sc.srno = nj.tran_id
      )
    `,
      { transaction: t },
    );

    await t.commit();
    return res.status(200).send({
      status: true,
      title: "Success",
      icon: "success",
      msg: "Application Shortlisted Successfully",
    });
  } catch (e) {
    // ❌ Rollback if anything fails
    if (t) await t.rollback();

    console.error("❌ shortlistcandidate error:", e);
    return res.status(500).send({
      status: false,
      title: "Error",
      icon: "error",
      message: e.message || "Failed to shortlist candidate",
    });
  } finally {
    if (sequelize) await sequelize.close();
  }
};





