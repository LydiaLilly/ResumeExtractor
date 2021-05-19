var express = require('express');
var router = express.Router();
var docx = require('../helper/docx')
const excelExport = require('../helper/excelExportBasic');
const resumeFolder = 'C:\\Users\\A742932\\Documents\\ReadResume\\Try\\extractresume\\resume\\'
const fs = require('fs');
var path = require('path');
var PdfReader = require('pdfreader').PdfReader;
const pdf = require('pdf-parse');


var skillSet = [/reactjs/i,/React.js/i,/react native/i, /java/i,/J2EE/i, /python/i, /Android/i, /iOS/i, /MySQL/i,/SQLite/i,/html/i,/automation/i,
  /Selenium/i,/Automated/i,/manual/i,
/BDD/i,/agile/i,/Jira/i,/TFS/i,/Azure/i,/Devops/i,/Data Warehousing/i,/Star schema/i,/EME/i,/UNIX/i,/Ab Initio/i,/ETL/i,
/Abinitio/i,/Oracle/i,/Data Warehousing/i,/UNIX/i,/Autosys/i,/ca7/i,/SQL/i,/Metadatahub/i,/.NET/i,/Shell/i,/Node/i,/NodeJS/i,
/Google Cloud/i,/Azure/i,/AWS/i,/Sales Force/i,/SAP/i,/Informatica/i,/Terrform/i,/Ansbile/i,/RestAPI/i,/SpringBoot/i,/Spring/i,
/Hibernate/i,/Microservices/i,/Performance Testing/i];

var testingSet = [/automation/i,/Selenium/i,/Automated/i,/Performance Testing/i]

var availableSkills = '';
var finalJsonObject = {
  'name':'',
  "skillSet": '',
  "contactNumber" : '',
  'yearOfExperience': ''  
};

const columnNames = ['Name','SkillSet', 'ContactNumber','Years of Experience','Profession'];
const headingColumnMap = {
    'Name':'name',
    'SkillSet': 'skillSet',
    'ContactNumber': 'contactNumber',
    'Years of Experience': 'yearOfExperience',
    'Profession':  'profession'
};

function generateJson(str = "",file){
  try{
 
    if((str && typeof(str) == "string") || file ){
      availableSkills = '';
      var fileName = file.replace(/.docx|.pdf/g, '');
      if (!fileName) 
        return;
  
        console.log("Name : " + fileName);

      finalJsonObject = { ...finalJsonObject, name: fileName};
       skillSet.map(skillRegex => {
         if(str.search(skillRegex) > -1) {
            availableSkills = availableSkills + ',' + skillRegex.source;
         }
      });
    
      finalJsonObject = { ...finalJsonObject, skillSet: availableSkills};

    
    
        var mobileNumber = str.match(/\d{12}|\d{10}/g);
        if(mobileNumber != null) {
          finalJsonObject = { ...finalJsonObject, contactNumber: mobileNumber[0]}; 
        } else {
          finalJsonObject = { ...finalJsonObject, contactNumber: ''}; 
        }
      
  
      var years = str.match(/([0-9.\+])+(\s?)+((y|Y)ears?)/g);
      if(years != null) {
        finalJsonObject = { ...finalJsonObject, yearOfExperience: years[0]};
      } else {
        finalJsonObject = { ...finalJsonObject, yearOfExperience: ''};
      }

      var professionType = 'Developer'
      var isAutomation = false;
      testingSet.map(testingRegex => {
        if(str.search(testingRegex) > -1) {
          professionType = 'Testing Automation';
          isAutomation = true;
        }
     });
      if(!isAutomation && str.search(/manual/i) > -1) {
        professionType = 'Manual Testing';
      } 
     
      finalJsonObject = { ...finalJsonObject, profession: professionType};
  
      return finalJsonObject;
     } 
  }catch(e){
     console.log(e);
  }

}

function generateExcel(finalArray){
    const workBook = excelExport.createSheet('Profile Details', columnNames, finalArray.sort((a,b)=> a - b), headingColumnMap);
    workBook.write('./skillDetail.xlsx');
}

/* GET users listing. */
router.get('/', async function(req, res, next) {
  res.send('Extract Profile Details');
  let output = [];
  const files = fs.readdirSync(resumeFolder);
  var i=0;
    for (const file of files) {
      if(path.extname(file)==".docx"){
        var docdata = await docx.extract(resumeFolder + file);
        output.push(generateJson(docdata,file));
      } else if(path.extname(file)==".pdf"){
        let dataBuffer = fs.readFileSync(resumeFolder + file);
        var pdfdata = await pdf(dataBuffer);      
        output.push(generateJson(pdfdata.text,file));
      }
      i++;
    }
  generateExcel(output);
});

module.exports = router;
