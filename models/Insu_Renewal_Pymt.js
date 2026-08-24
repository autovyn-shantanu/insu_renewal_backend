const Sequelize = require("sequelize");
const { DataTypes } = require("sequelize");
const Joi = require("joi");

const Insu_Renewal_Pymt = function (sequelize, DataTypes) {
  return sequelize.define(
    "InsuRenewalPymt",
    {
      UTD: {
        type: DataTypes.INTEGER,
        allowNull: false,
        primaryKey: true,
        autoIncrement: true,
      },

      INSU_TYPE: {
        type: DataTypes.INTEGER,
        allowNull: false,
      },

      PREMIUM_AMOUNT: {
        type: DataTypes.DECIMAL(18, 2),
        allowNull: true,
      },

      PYMT_MODE: {
        type: DataTypes.STRING(100),
        allowNull: false,
      },

      // ✅ only once
      PYMT_DATE: {
        type: DataTypes.DATEONLY,
        allowNull: false,
      },

      PYMT_AMOUNT: {
        type: DataTypes.DECIMAL(18, 2),
        allowNull: true,
      },

      PYMT_CODE: {
        type: DataTypes.STRING(100),
        allowNull: true,
      },

      PYMT_REMARK: {
        type: DataTypes.STRING(100),
        allowNull: true,
      },

      PYMT_STATUS: {
        type: DataTypes.INTEGER,
        allowNull: true,
      },

      ACNT_APPR_CODE: {
        type: DataTypes.STRING(100),
        allowNull: true,
      },

      ACNT_APPR_REMARK: {
        type: DataTypes.STRING(100),
        allowNull: true,
      },

      ACNT_APPR_STATUS: {
        type: DataTypes.INTEGER,
        allowNull: true,
      },

      ACNT_APPR_DATE: {
        type: DataTypes.DATEONLY,
        allowNull: true,
      },

      POLICY_GENERATE: {
        type: DataTypes.INTEGER,
        allowNull: true,
      },

      UTR: {
        type: DataTypes.INTEGER, // (agar UTR alpha-numeric hota hai to VARCHAR better)
        allowNull: true,
      },

      BANK_NAME: {
        type: DataTypes.STRING(100),
        allowNull: true,
      },

      REMARKS: {
        type: DataTypes.STRING(100),
        allowNull: true,
      },

      TRAN_ID: {
        type: DataTypes.INTEGER,
        allowNull: true,
      },

      DOC_PATH: {
        type: DataTypes.STRING(100),
        allowNull: true,
      },

      CREATED_AT: {
        type: DataTypes.DATE,
        allowNull: false,
        // DB default GETDATE() hai, phir bhi safe
        defaultValue: Sequelize.literal("GETDATE()"),
      },

      // Temporal columns (GENERATED ALWAYS) -> insert/update me mat bhejna
      VALIDFROM: {
        type: DataTypes.DATE,
        allowNull: false,
      },
      VALIDTO: {
        type: DataTypes.DATE,
        allowNull: false,
      },
    },
    {
      sequelize,
      tableName: "INSU_RENEWAL_PYMT",
      schema: "dbo",
      timestamps: false,
      indexes: [
        { name: "PK_INSU_RENEWAL_PYMT", unique: true, fields: [{ name: "UTD" }] },
        { name: "IX_INSU_RENEWAL_PYMT_TRAN_ID", fields: [{ name: "TRAN_ID" }] },
        { name: "IX_INSU_RENEWAL_PYMT_PYMT_DATE", fields: [{ name: "PYMT_DATE" }] },
      ],
    }
  );
};

// Joi schema (API validation/import validation)
const insuRenewalPymtSchema = Joi.object({
  UTD: Joi.number().integer().positive().optional(),

  INSU_TYPE: Joi.number().integer().required(),

  PREMIUM_AMOUNT: Joi.number().precision(2).allow(null, "").optional(),

  PYMT_MODE: Joi.string().max(100).required(),

  PYMT_DATE: Joi.date().raw().required(),

  PYMT_AMOUNT: Joi.number().precision(2).allow(null, "").optional(),

  PYMT_EMPCODE: Joi.string().max(100).allow(null, "").optional(),
  PYMT_REMARK: Joi.string().max(100).allow(null, "").optional(),

  PYMT_STATUS: Joi.number().integer().allow(null, "").optional(),

  ACNT_APPR_CODE: Joi.string().max(100).allow(null, "").optional(),
  ACNT_APPR_REMARK: Joi.string().max(100).allow(null, "").optional(),
  ACNT_APPR_STATUS: Joi.number().integer().allow(null, "").optional(),
  ACNT_APPR_DATE: Joi.date().raw().allow(null, "").optional(),

  POLICY_GENERATE: Joi.number().integer().allow(null, "").optional(),

  UTR: Joi.alternatives()
    .try(Joi.number().integer(), Joi.string().max(50))
    .allow(null, "")
    .optional(),

  BANK_NAME: Joi.string().max(100).allow(null, "").optional(),
  REMARKS: Joi.string().max(100).allow(null, "").optional(),

  TRAN_ID: Joi.number().integer().allow(null, "").optional(),

  DOC_PATH: Joi.string().max(260).allow(null, "").optional(), // DB 100 hai, but URL 260 tak ho sakta hai

  CREATED_AT: Joi.date().raw().allow(null, "").optional(),

  // Temporal columns -> always strip
  VALIDFROM: Joi.any().strip(),
  VALIDTO: Joi.any().strip(),
});

module.exports = { Insu_Renewal_Pymt, insuRenewalPymtSchema };