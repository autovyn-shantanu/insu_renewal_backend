module.exports = (sequelize, DataTypes) => {
  return sequelize.define(
    "FOLLOWUP_DETAIL",
    {
      TRAN_ID: { type: DataTypes.INTEGER, allowNull: false, primaryKey: true },

      FOLLOWUP_STATUS: { type: DataTypes.STRING(100), allowNull: true },
      FOLLOWUP_DATE: { type: DataTypes.DATEONLY, allowNull: true },

      LAST_FOLLOWUP_DATE: {
        type: DataTypes.DATEONLY,
        allowNull: true,
        field: "LAST_FOLLOWUP_DATE",
      },
     FOLLOWUP_TIME: {
        type: DataTypes.TIME,     // MSSQL time
        allowNull: true,
        field: "FOLLOWUP_TIME",
      },

      CALL_ID: { type: DataTypes.STRING(100), allowNull: true, field: "call_id" },
     EXPORT_TYPE: { type: DataTypes.INTEGER, allowNull: false, defaultValue: 1 }, // ✅ add this


      REMARKS: { type: DataTypes.STRING(100), allowNull: true },
    },
    {
      tableName: "FOLLOWUP_DETAILS",
      timestamps: false,
    }
  );
};