const express = require("express");
const Crm = require("../routes/Crm");
const multer = require("multer");
const router = express.Router();

const excelUpload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024, // ✅ fileSize use karo (fieldSize nahi)
  },
}).fields([{ name: "excel", maxCount: 1 }]);

const paymentUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB
});


router.get("/sampel", Crm.downloadInsuranceRenualSampleExcel); // ✅ yaha multer ki need nahi
router.post("/import", excelUpload, Crm.importInsuRenewalExcel); // ✅ yaha multer zaroori hai
router.post("/filter", Crm.getInsuRenewalByDateRange);
router.post("/reminders", Crm.getInsuRenewalReminders);
router.post('/followup',Crm.saveInsuRenewalFollowupByVehicle)
router.post("/history", Crm.getFollowupHistoryByVehicle);
router.post("/getVehicleByRegNo", Crm.getVehicleByRegNo)
router.post("/getInsuranceAndPaymentDropdowns", Crm.getInsuranceAndPaymentDropdowns);
router.post( "/SaveInsuranceRenewal",paymentUpload.single("PaymentProof"),Crm.SaveInsuranceRenewal );
router.post("/getAll",  Crm.getAllInsuranceRenewals);   // ← NEW
router.post("/getOne",  Crm.getInsuranceRenewalById)
router.post("/approval",Crm.updateInsuranceRenewalApproval)
router.post("/getApprovedOne", Crm.getAllApprovedInsuranceRenewals);
router.post('/makecall',Crm.makeInsuranceRenewalCall)
router.post("/webhook", Crm.callmaticWebhook);
// routes/excel.js (example)
router.post("/calling/customers", Crm.getCallingCustomers);
router.post("/calling/history", Crm.getCallingHistoryByMobile);
router.get("/calling/recording/:callId", Crm.streamCallRecordingByCallId);
router.get("/insucalling/get", Crm.getInsuCallingConfig);
router.post("/insucalling/save", Crm.createInsuCallingConfig);
router.put("/insucalling/update", Crm.updateInsuCallingConfig);
router.patch("/insucalling/toggle", Crm.toggleInsuCallingConfigStatus);
router.post('/dashboard',Crm.getInsuranceDashboardMetrics)
router.post('/od',Crm.OD_report); 
router.post('/bonvoice',express.raw({type: 'application/json'}),Crm.bonvoiceWebhook)
router.post('/attendance',Crm.attendance_pivot_report)
router.post("/getPendingTasksByExecutive",   Crm.getPendingInsuranceByExecutive);
router.post("/getEmployees",Crm.getInsuranceExecutives)
router.post("/transferInsuranceWorkload", Crm.transferInsuranceWorkload);
router.post('/aniversery/today',Crm.getEmployeeAnniversaries)
router.post('/assignwork',Crm.assignInsuranceRenewalsToDSE)
// router.post('/getExpiringInsuranceRenewals',excel.getExpiringInsuranceRenewals)
router.post('/insu/whatsapp',Crm.SendInsuranceRenewalWhatsAppToCustomer)
router.get('/InsuranceRenewalCustomerView', Crm.InsuranceRenewalCustomerView);
router.post('/SaveInsuranceCustomerResponse', Crm.SaveInsuranceCustomerResponse);
// router.post('/getDSEOwnTasks', excel.getDSEOwnTasks);
// router.post('/getReportingManagerTeamTasks', excel.getReportingManagerTeamTasks);



module.exports = router;