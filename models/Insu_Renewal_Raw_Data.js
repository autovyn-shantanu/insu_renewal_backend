// models/Insu_Renewal_Raw_Data.js
const Sequelize = require("sequelize");

const Insu_Renewal_Raw_Data = (sequelize, DataTypes) => {
  return sequelize.define(
    "InsuRenewalRawData",
    {
      UTD: {
        type: DataTypes.INTEGER,
        allowNull: false,
        primaryKey: true,
        autoIncrement: true,
      },

      CUST_NAME: {
        type: DataTypes.STRING(100), // nvarchar(100)
        allowNull: true,             // ✅ ALTER के बाद NULL allowed
      },

      CUST_MOB_NO: {
        type: DataTypes.BIGINT,      // bigint
        allowNull: true,
      },

      POLICY_NAME: {
        type: DataTypes.STRING(100), // nvarchar(100)
        allowNull: true,             // ✅
      },

      POLICY_NUMBER: {
        type: DataTypes.STRING(100), // varchar(100)
        allowNull: true,             // ✅
      },

      VEHICAL_REG_NO: {
        type: DataTypes.STRING(20),  // varchar(20)
        allowNull: true,
      },

      MODEL_NAME: {
        type: DataTypes.STRING(100), // nvarchar(100)
        allowNull: true,
      },

      POLICY_START_DATE: {
        type: DataTypes.DATEONLY,    // date
        allowNull: true,             // ✅
      },

      POLICY_END_DATE: {
        type: DataTypes.DATEONLY,    // date
        allowNull: true,             // ✅
      },

      DSC_EMPCODE: {
        type: DataTypes.STRING(50),  // varchar(50)  ✅ exact length
        allowNull: true,
      },

      DSC_NAME: {
        type: DataTypes.STRING(100), // varchar(100)
        allowNull: true,
      },

      DSC_MOB_NO: {
        type: DataTypes.BIGINT,      // ✅ INT से BIGINT किया (overflow fix)
        allowNull: true,
      },

      CREATED_AT: {
        type: DataTypes.DATE,        // datetime
        allowNull: false,
        defaultValue: Sequelize.fn("GETDATE"),
      },
    },
    {
      tableName: "INSU_RENEWAL_RAW_DATA",
      schema: "dbo",
      timestamps: false,
    }
  );
};

module.exports = { Insu_Renewal_Raw_Data };