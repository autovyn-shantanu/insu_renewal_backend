const template=require("../routes/template.js")
const express= require("express")
const router=express.Router()

router.post("/AllSelectempdata",template.AllSelectempdata)
module.exports=router