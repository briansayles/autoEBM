 //
 // Call this from the terminal with 1 required argument and 4 optional argument flags
 //   arg = dataFileName (string) = the relative path to the data file (JSON) containing entries to be processed
 //   -noExcel (boolean flag): setting this flag suppresses the creation of an Excel file summarizing the results
 //   -noLabels (boolean flag): setting this flag suppresses the creation of Word files for each label
 //   -noMergeFile (boolean flag): setting this flag suppresses the creation of a Word file compilation of all of the label files using the label_quantity field in the data file
 //   jobNumber (string) = job number to be used on labels and output filename
 // eg: node labelMaker.js dataFileName='/customers/template customer/input/testEntries.xlsx' -noExcel -noLabels -noMerge
 // typical usage: node makeLabels.js dataFileName='customers/___/input/___.xlsx' jobnumber='99999999'
 //

import { saveToExcel, generateMailMergeDOCX } from './utilityFunctions.js';
import { __dirname, __filename } from './config.js';
import path from 'path';
import XLSX from 'xlsx';

makeLabels();

function readAFDataFile(excelFilename) {
  const afDataBook = XLSX.readFile(`./${excelFilename}`);
  console.log('found workbook'); 
  const afDataSheet=afDataBook.Sheets['AF Labels Data'];
  console.log('found worksheet');
  const afJSONData=JSON.stringify(XLSX.utils.sheet_to_json(afDataSheet));
  console.log('converted to JSON');
  return afJSONData;
}

async function makeLabels() {
  // Function to create lables from a template file and a data file, and to create a merged file if desired, and a summary Excel file if desired
  const args = process.argv.slice(2).reduce((acc, arg) => {
    let [k, v = true] = arg.split('=')
    acc[k] = v
    return acc
  }, {})
  console.log(args);
  var start = new Date().getTime();
  console.log('attempting to read from ' + args.dataFileName);
  const afEntries = JSON.parse(readAFDataFile(args.dataFileName));
  const jobNumber = args.jobNumber || "";
  const customerPath = path.dirname(path.dirname(path.resolve(__dirname, args.dataFileName)));
  const customerDataFileName = `${customerPath}/customer_data/customerData.json`;
  const customerDataJSON = await import(customerDataFileName, {with: {type: 'json'}})
  const customerData = customerDataJSON.default[0].customer;
  const ieArray = [];
  customerData.ie_breakpoints.forEach(ie_breakpoint => ieArray.push(ie_breakpoint.calories));
  var end = new Date().getTime();
  var time = end - start;
  console.log(`Retrieved ${afEntries.length} AF entries for processing in ${time/1000} seconds`);
  start = new Date().getTime();
  const outputVariables = [];
  const excelOutputs = [];
  const sources = customerData.sources;
  afEntries.forEach((afEntry, afEntryIndex, afEntriesArray) => {
      outputVariables.push(
        {
          ...afEntry,
          timestamp: new Date().toISOString(),
          varAFB: afEntry["Arc Flash Boundary (in)"],
          varAFBFeetInches: afEntry["Arc Flash Boundary (ft-in)"],
          varVoltage: afEntry["Voltage (kV)"]*1000,
          varkV: afEntry["Voltage (kV)"],
          varRAB: afEntry["Restricted Approach Boundary (in)"],
          varRABFeetInches: afEntry["Restricted Approach Boundary (ft-in)"],
          varLAB: afEntry["Limited Approach Boundary (in)"],
          varLABFeetInches: afEntry["Limited Approach Boundary (ft-in)"],
          varEquipmentName: afEntry["ID"],
          varFedFrom: afEntry["Source Protective Device"],
          varEquipmentLocation: afEntry["Equipment Location"] || afEntry["Job Name"] || "",
          varPPE: afEntry["PPE Level (Site-Specific)"],
          varMaxIE: afEntry["Incident Energy (cal/cm2)"],
          varWorkingDistance: afEntry["Working Distance (in)"],
          varWorkingDistanceFeetInches: afEntry["Working Distance (ft-in)"],
          varQuantity: afEntry["Label Quantity"] || 1,
          varJobNumber: afEntry["Job Number"]
        }
      );
      if (!args['-noExcel']) {
        excelOutputs.push(
          {
            ...afEntry,
            timestamp: new Date().toISOString(),
          }
        );
      }
    // },);
  },);

  if (!args['-noExcel']) {
    try {
      saveToExcel(excelOutputs, customerData, customerPath, jobNumber);
    } catch (err) {
      console.log(err.message);
    }
  }

  if (!args['-noLabels']) {
    try {
      generateMailMergeDOCX(outputVariables, customerData, customerPath, args['-noMerge'], jobNumber);
    } catch (err) {
      console.log(err.message);
    }
  }
  end = new Date().getTime();
  time = end - start;
  console.log(`Process took ${time/1000} seconds for ${afEntries.length} items`);
}