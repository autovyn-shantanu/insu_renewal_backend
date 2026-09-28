const { dbname } = require("../utils/dbconfig");

exports.AllSelectempdata = async function (req, res) {
  const sequelize = await dbname(req, req.headers.compcode);
  const empcodes = req.body.empcode
    ? (Array.isArray(req.body.empcode) ? req.body.empcode : [req.body.empcode]).filter(Boolean)
    : [];

  try {
    const data = await sequelize.query(
      `
      SELECT
        ISNULL(e.EMPCODE, '') AS EMPCODE,
        ISNULL(CONCAT(e.EMPFIRSTNAME, ' ', e.EMPLASTNAME), '') AS EMPNAME,
        ISNULL(e.EMPLOYEEDESIGNATION, '') AS EMPDESIGNATION,
        ISNULL(e.EMPLOYEEDESIGNATION, '') AS EMPLOYEEDESIGNATION,
        pt.APPR_1_STAT,
        CASE
          WHEN pt.APPR_1_STAT = 1 THEN 'Approved'
          WHEN pt.APPR_1_STAT = 0 THEN 'Reject'
          ELSE 'Pending'
        END AS STATUS
      FROM EMPLOYEEMASTER e
      LEFT JOIN PRINT_TEMPLATES pt ON e.EMPCODE = pt.EMPCODE
      ${empcodes.length > 0 ? "WHERE e.EMPCODE IN (:empcodes)" : ""}
      `,
      {
        replacements: empcodes.length > 0 ? { empcodes } : {},
        type: sequelize.QueryTypes.SELECT,
      }
    );

    res.status(200).send({ success: true, data: data });
  } catch (e) {
    console.log(e);
    res.status(500).send({ success: false, message: e.message });
  } finally {
    if (sequelize) {
      await sequelize.close();
    }
  }
};
