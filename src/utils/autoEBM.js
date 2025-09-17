 //
 // Call this from the server.js with 5 arguments
 //   dataFileName (string) = the path to the data file (JSON) containing entries to be processed
 //   noExcel (string true | false): setting this flag suppresses the creation of an Excel file summarizing the results
 //   noLabels (string true | false): setting this flag suppresses the creation of Word files for each label
 //   noMergeFile (string true | false): setting this flag suppresses the creation of a Word file compilation of all of the label files using the label_quantity field in the data file
 //   jobNumber (string) = job number to be used on labels and output filename
 //

import { __dirname, __filename } from '../config.js';
import ebmStaticValues from '../static_data/ebmStaticValues.json' with {type: 'json'};
import shockBoundaries from '../static_data/shockBoundaries.json' with {type: 'json'};
import arcFlashBoundaries from '../static_data/arcFlashBoundaries.json' with {type: 'json'};
import XLSX from 'xlsx';
import PizZip from 'pizzip';
import fs from 'fs';
import Docxtemplater from 'docxtemplater';
import expressionParser from "docxtemplater/expressions.js";
import DocxMerger from 'docx-merger';
import archiver from 'archiver';

export async function applyEnergyBoundaryMethod({dataFileName, noExcel, noLabels, noMerge, jobNumber, configFileName, templateFileName} = {}) {
  const start0 = new Date().getTime();
  var start = new Date().getTime();
  const createExcel = noExcel == 'false'
  const createIndividualLabels = noLabels == 'false';
  const createMergeFile = noMerge == 'false';
  jobNumber = jobNumber || "";
  ebmStaticValues.sort((a, b)=> b.kA - a.kA);
  const ebmEntries = readEnergyBoundaryEntriesFromXLSX(dataFileName);
  // console.log(`${__dirname}`)
  // const customerPath = `${__dirname}/customers/${customerName}`;
  // console.log(customerPath);
  // const customerDataFileName = `${customerPath}/customer_data/customerData.json`;
  // const customerDataJSON = await import(customerDataFileName, {with: {type: 'json'}})
  const customerDataJSON = await import(configFileName, {with: {type: 'json'}});
  if (!customerDataJSON || !customerDataJSON.default || customerDataJSON.default.length == 0) {
    throw new Error(`Configuration file ${configFileName} is not valid. Please check the customer configuration.`);
  }
  // const customerTemplateFileName = `${customerPath}/customer_data/${customerDataJSON.default[0].customer.template}`;
  // console.log(customerTemplateFileName);
  // if (!fs.existsSync(customerTemplateFileName)) {
  //   throw new Error(`Template file ${customerTemplateFileName} does not exist. Please check the customer configuration.`);
  // }
  // console.log(`Using template file: ${customerTemplateFileName}`);
  const templateData = fs.readFileSync(templateFileName, "binary");


  const outputVariables = [];
  const excelOutputs = [];
  const customerData = customerDataJSON.default[0].customer;
  const sources = customerData.sources;
  const ieArray = [];
  customerData.ie_breakpoints.forEach(ie_breakpoint => ieArray.push(ie_breakpoint.calories));
  try {
      ebmEntries.end_use_equipment.forEach((equipmentItem, equipmentIndex, equipmentArray) => {
        // console.log(`Processing item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}`);
        // console.log(` - OCPD: ${equipmentItem.ocpd.amps}A ${equipmentItem.ocpd.type} (${equipmentItem.ocpd.class})`);
        // console.log(` - Distance: ${convertToNumber(equipmentItem.distance_ft)} ft`);
        // console.log(` - Source: ${equipmentItem.source}`);
        // console.log(` - Location: ${equipmentItem.location}`);
        // console.log(` - Label Quantity: ${equipmentItem.label_quantity}`);
        let recommendation = "";
        let recommendRK1 = false;
        let equipmentPPELevelRK1 = "";
        const source = sources.filter((source)=> source.name == equipmentItem.source)[0];
        if (!source) {
          // console.log(` - No source found matching name ${equipmentItem.source}`);
          throw new Error(`Error processing item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}. No source found matching name ${equipmentItem.source}. Please check the data file and customer data.`);
        }
        // console.log(` - Source found: ${source.name} (${source.voltage} V, ${source.kA} kA)`);
        const ebmStaticLine = ebmStaticValues.filter((ebmStaticValue) => 
          (ebmStaticValue.kA <= source.kA) * (JSON.stringify(ebmStaticValue.ocpd) === JSON.stringify(equipmentItem.ocpd)) * (ebmStaticValue.voltage == source.voltage) 
        )[0];
        if (!ebmStaticLine) {
          // console.log(` - No EBM static line found for source ${source.name} with ${source.kA} kA, ${source.voltage} V, and OCPD ${equipmentItem.ocpd.amps}A ${equipmentItem.ocpd.type} (${equipmentItem.ocpd.class})`);
          throw new Error(`Error processing item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}. No EBM data found for source ${source.name} with ${source.kA} kA, ${source.voltage} V, and OCPD ${equipmentItem.ocpd.amps}A ${equipmentItem.ocpd.type} (${equipmentItem.ocpd.class}). Please check the data file and customer data.`);
        }
        // console.log(` - EBM static line found: ${ebmStaticLine.kA} kA, ${ebmStaticLine.voltage} V, ${ebmStaticLine.ocpd.amps}A ${ebmStaticLine.ocpd.type} (${ebmStaticLine.ocpd.class})`);
        const equipmentWorkingDistance = ebmStaticLine.working_distance_in;
        const ieBreakPoints = ebmStaticLine.boundaries.filter((boundary) =>  ieArray.includes(boundary.calories));
        const equipmentIEBreakpoint = ieBreakPoints.filter(breakpoint => convertToNumber(breakpoint.distance_ft) >= convertToNumber(equipmentItem.distance_ft))[0]
        const equipmentPPELevel = customerData.ie_breakpoints.filter(ie_breakpoint => ie_breakpoint.calories == equipmentIEBreakpoint.calories)[0].name;
        // console.log(`  - Equipment IE Breakpoint: ${equipmentIEBreakpoint.calories} cal/cm2`);
        // console.log(`  - Equipment Required PPE Level: ${equipmentPPELevel}`);
        const equipmentMaxIE = equipmentIEBreakpoint.calories;
        if (equipmentMaxIE > ieBreakPoints[0].calories) {
          if(equipmentItem.ocpd.class == "RK5") {
            recommendation = `Warning: Calculated max IE of ${equipmentMaxIE} cal/cm2 exceeds customer's minimum PPE level of ${ieBreakPoints[0].calories} cal/cm2.`; 
            const ebmStaticLineRK1 = ebmStaticValues.filter((ebmStaticValue) => 
              (ebmStaticValue.kA <= source.kA) * (JSON.stringify(ebmStaticValue.ocpd) === JSON.stringify({...equipmentItem.ocpd, class: "RK1"})) * (ebmStaticValue.voltage == source.voltage)
            )[0];
            const ieBreakPointsRK1 = ebmStaticLineRK1.boundaries.filter((boundary) =>  ieArray.includes(boundary.calories));
            const equipmentIEBreakpointRK1 = ieBreakPointsRK1.filter(breakpoint => convertToNumber(breakpoint.distance_ft) >= convertToNumber(equipmentItem.distance_ft))[0]
            const equipmentMaxIERK1 = equipmentIEBreakpointRK1.calories;
            equipmentPPELevelRK1 = customerData.ie_breakpoints.filter(ie_breakpoint => ie_breakpoint.calories == equipmentIEBreakpointRK1.calories)[0].name;
            if (equipmentMaxIERK1 < equipmentMaxIE) {
              recommendation += ` Using a class RK1 fuse will reduce the required PPE level from ${equipmentPPELevel} to ${customerData.ie_breakpoints.filter(ie_breakpoint => ie_breakpoint.calories == equipmentMaxIERK1)[0].name}.`;
              recommendRK1 = true;
            } else {
              recommendation += ` Using a class RK1 will not reduce the required PPE level from ${equipmentPPELevel}.`;
              recommendRK1 = false;
              equipmentPPELevelRK1 = "";
            } 
          } else if (equipmentItem.ocpd.class == "RK1") {
            recommendation = `Warning: Calculated max IE of ${equipmentMaxIE} cal/cm2 exceeds customer's minimum PPE level of ${ieBreakPoints[0].calories} cal/cm2. Consider using a class RK1 fuse with a lower amp rating, feeding from another source, or reducing distance.`;
            equipmentPPELevelRK1 = "";
            recommendRK1 = false;
          }
        } else {
          recommendation = "";
          equipmentPPELevelRK1 = "";
        }
        // console.log(`  - Recommendation: ${recommendation == "" ? "None" : recommendation}`);
        // console.log(`  - Recommend RK1 Fuse? ${recommendRK1 ? "Yes" : "No"}`);
        // console.log(`  - Equipment PPE Level with RK1 Fuse: ${equipmentPPELevelRK1 == "" ? "N/A" : equipmentPPELevelRK1}`);
        const equipmentShockBoundaries = shockBoundaries.filter(shockBoundary => shockBoundary.voltage_max >= source.voltage)[0];
        const equipmentLimitedApproachBoundaryInches = equipmentShockBoundaries.limited_approach_in;
        const equipmentRestrictedApproachBoundaryInches = equipmentShockBoundaries.restricted_approach_in;
        const equipmentArcFlashBoundaryInches = arcFlashBoundaries.filter(
          arcFlashBoundary => (arcFlashBoundary.voltage_max >= source.voltage) * (arcFlashBoundary.equipment_type == "Equipment"))[0].boundaries.filter(bndry => bndry.calories == equipmentIEBreakpoint.calories)[0].distance_in;
        outputVariables.push(
          {
            dataProvided: equipmentItem,
            timestamp: new Date().toISOString(),
            varAFB: equipmentArcFlashBoundaryInches,
            varAFBFeetInches: toFeetInches(equipmentArcFlashBoundaryInches),
            varVoltage: source.voltage,
            varkV: source.voltage/1000,
            varRAB: equipmentRestrictedApproachBoundaryInches,
            varRABFeetInches: toFeetInches(equipmentRestrictedApproachBoundaryInches),
            varLAB: equipmentLimitedApproachBoundaryInches,
            varLABFeetInches: toFeetInches(equipmentLimitedApproachBoundaryInches),
            varEquipmentName: equipmentItem.name,
            varFedFrom: source.name,
            varEquipmentLocation: equipmentItem.location,
            varPPE: equipmentPPELevel,
            varMaxIE: equipmentMaxIE,
            varWorkingDistance: equipmentWorkingDistance,
            varWorkingDistanceFeetInches: toFeetInches(equipmentWorkingDistance),
            varQuantity: equipmentItem.label_quantity,
            varJobNumber: jobNumber
          }
        );
        if (createExcel) {
          excelOutputs.push(
          {
            "Equipment Name": equipmentItem.name,
            "Source": equipmentItem.source,
            "OCPD": `${equipmentItem.ocpd.amps}A ${equipmentItem.ocpd.type} (${equipmentItem.ocpd.class})`,
            "Circuit Length (ft)": equipmentItem.distance_ft,
            "Location": equipmentItem.location,
            "Label Quantity": equipmentItem.label_quantity,
            "Arc Flash PPE Level": equipmentPPELevel,
            "Arc Flash Max IE at Working Distance (cal/cm2)": equipmentMaxIE,
            "Working Distance (in)": equipmentWorkingDistance,
            "Working Distance (ft-in)": toFeetInches(equipmentWorkingDistance),
            "Arc Flash Boundary (in)": equipmentArcFlashBoundaryInches,
            "Arc Flash Boundary (ft-in)": toFeetInches(equipmentArcFlashBoundaryInches),
            "Voltage (V)": source.voltage,
            "Restricted Approach Boundary (in)": equipmentRestrictedApproachBoundaryInches,
            "Restricted Approach Boundary (ft-in)": toFeetInches(equipmentRestrictedApproachBoundaryInches),
            "Limited Approach Boundary (in)": equipmentLimitedApproachBoundaryInches,
            "Limited Approach Boundary (ft-in)": toFeetInches(equipmentLimitedApproachBoundaryInches),
            "Timestamp": new Date().toISOString(),
            "Job Number": jobNumber,
            "Recommend RK1 Fuse?": recommendRK1 ? "Yes" : "No",
            "Equipment PPE Level with RK1 Fuse": equipmentPPELevelRK1,
            "Recommendation": recommendation
          });
        }
      },);
  } catch (error) {
      return {
        message: error.message,
        error: true,
        excelFilePath: null,
        wordFilePath: null,
        labelsZipPath: null
      };    
  }
  var end = new Date().getTime();
  var time = end - start;
  start = new Date().getTime();
  let excelResult;
  let wordResult;  
  const finishTimestamp = '(' + new Date().toISOString() + ')';
  if (createExcel) {
    try {
      excelResult = await saveToExcel(excelOutputs, customerData, jobNumber, finishTimestamp);
      // console.log(`Excel file created at: ${excelResult}`);
    } catch (err) {
      console.log(err.message);
    }
  }
  if (createIndividualLabels) {
    try {
      wordResult = await generateMailMergeDOCX(outputVariables, customerData, createMergeFile, jobNumber, finishTimestamp, templateData);
    } catch (err) {
      console.log(err.message);
    }
  }
  if (createIndividualLabels) {
    try {
      const zipPath = await createLabelsZip(jobNumber, finishTimestamp);
      // console.log(`Labels zip file created at: ${zipPath}`);
      wordResult.labelsZipPath = zipPath;
    } catch (err) {
      console.log(err.message);
    }
  }
  end = new Date().getTime();
  time = end - start0;
  return {
    message: `autoEBM processing complete for ${ebmEntries.end_use_equipment.length} entries in ${time/1000} seconds`, 
    error: false,
    excelFilePath: excelResult,
    wordFilePath: createMergeFile ? wordResult.mergeFilePath : null,
    labelsZipPath: createIndividualLabels ? wordResult.labelsZipPath : null
  };
}

async function createLabelsZip(jobNumber, finishTimestamp) {
  const labelsDir = `./output/${finishTimestamp}/individual labels`;
  // const finishTimestamp = '(' + new Date().toISOString() + ')';
  const zipPath = `./output/${finishTimestamp}/individual_labels ${jobNumber !== "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}.zip`;

  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => {
      console.log(`Created zip file at: ${zipPath} (${archive.pointer()} total bytes)`);
      resolve(zipPath);
    });

    archive.on('error', (err) => {
      reject(err);
    });

    archive.pipe(output);
    archive.directory(labelsDir, false);
    archive.finalize();
  });
}

function toFilenameFriendlyFormat(input) {
  return input
    .trim()
    .replace(/[^a-zA-Z0-9() ]/g, '_')
    .replace(/-+/g, '-');
}

function toFeetInches(inches) {
  return (
    Math.floor(inches / 12) + "' " + (inches % 12) + '"'
  );
}

function convertToNumber(input) {
  // Remove any non-digit characters (like "st", "nd", "rd", "th")
  if (isNaN(input)) { 
    const number = parseInt(input.replace(/(st|nd|rd|th)$/i, ''), 10);
    return isNaN(number) ? null : number; // Return null if conversion fails
  } else if (typeof input === 'number') {
    return input;
  }
}

async function saveToExcel(excelOutputs, customer, jobNumber, finishTimestamp) {
  const start = new Date().getTime();
  const worksheet = XLSX.utils.json_to_sheet(excelOutputs);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'AF Results');
  fs.mkdirSync(`./output/${finishTimestamp}`, { recursive: true });
  // const finishTimestamp = '(' + new Date().toISOString() + ')';
  const excelFilename = toFilenameFriendlyFormat(`${customer.name} AF Results ${jobNumber !== "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}`);
  const filePath = `./output/${finishTimestamp}/${excelFilename}.xlsx`;
  XLSX.writeFile(workbook, filePath);
  const end = new Date().getTime();
  const time = end - start;
  console.log(`Save to Excel took ${time / 1000} seconds for ${excelOutputs.length} items.`);
  return filePath;
}

async function generateMailMergeDOCX(data, customer, createMergeFile, jobNumber, finishTimestamp, templateFile) {
  var start = new Date().getTime();
  fs.mkdirSync(`./output/${finishTimestamp}/individual labels/`, { recursive: true });
  // const content = fs.readFileSync(`${customerPath}/customer_data/${customer.template}`, "binary");
  var docxFiles = [];
  data.forEach((item, index, array) => {
    if (item.varQuantity > 0) {
      // var timestamp = '(' + new Date().toISOString().slice(0, 10) + ')';
      var filename = toFilenameFriendlyFormat(`${item.varEquipmentName} ${jobNumber != "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}`);
      item.varLabelID = `${item.varEquipmentName}-${finishTimestamp}`;
      // const zip = new PizZip(content);
      const zip = new PizZip(templateFile);
      var doc = new Docxtemplater(zip, {
        parser: expressionParser,
        linebreaks: true,
        paragraphLoop: true,
      });
      doc.render(item);
      var buffer = doc.toBuffer();
      try {
        fs.writeFileSync(`./output/${finishTimestamp}/individual labels/${filename}.docx`, buffer, []);
        if (createMergeFile) {
          for (let i = 0; i < item.varQuantity; i++) {
            docxFiles.push(buffer);
          }
        }
      }
      catch (err) {
        console.error("Error: ", err.message);
      }
    }
  });
  var end = new Date().getTime();
  var time = end - start;
  console.log(`Creation of individual label files took ${time / 1000} seconds for ${data.length} items`);

  let mergeFilename = '';
  if (createMergeFile && docxFiles.length > 0) {
    start = new Date().getTime();
    const docxMerger = new DocxMerger({}, docxFiles);
    // const finishTimestamp = '(' + new Date().toISOString() + ')';
    mergeFilename = toFilenameFriendlyFormat(`${customer.name} AF Labels ${jobNumber != "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}`);
    docxMerger.save('nodebuffer', (data) => {
      fs.writeFileSync(`./output/${finishTimestamp}/${mergeFilename}.docx`, data, (err) => {
        console.log(err.message);
      });
    });
    end = new Date().getTime();
    time = end - start;
    console.log(`Mail merge execution took ${time / 1000} seconds for ${docxFiles.length} pages`);
  }
  return (
    {
      message: 'Mail merge complete',
      mergeFilePath: createMergeFile ? `./output/${finishTimestamp}/${mergeFilename}.docx` : 'No merged file created'
    }
  )
}

function readEnergyBoundaryEntriesFromXLSX(excelFilename) {
  const ebmEntryBook = XLSX.readFile(`${excelFilename}`);  
  const ebmEntrySheet = ebmEntryBook.Sheets['Order Form'];
  const ebmEntriesJSON = XLSX.utils.sheet_to_json(ebmEntrySheet);
  // console.log(`Read ${ebmEntriesJSON.length} entries from data file ${excelFilename}`);
  // console.log(ebmEntriesJSON);
  const ebmJSONData = [];
  ebmEntriesJSON.forEach((entry) => {
    try {
        const name = entry['Title (Equipment Name)'];
        const distance_ft = entry['Circuit Length (ft)'];
        const source = entry['Source'];
        const location = entry['Equipment Location (Columns)'];
        const ampsSymbolIndex = entry['OCPD'].indexOf("A ");
        const ocpdAmps = parseInt(entry['OCPD'].slice(0, ampsSymbolIndex)) || 0;
        const ocpdType = entry['OCPD'].indexOf("Fuse") !== -1 ? "Fuse" : 
          entry['OCPD'].indexOf("MCCB") !== -1 ? "MCCB" :
          entry['OCPD'].indexOf("Force") !== -1 ? "FORCE" : "N/A";
        const ocpdClass = entry['OCPD'].indexOf("Class RK5") !== -1 ? "RK5" :
          entry['OCPD'].indexOf("Class RK1") !== -1 ? "RK1" : 
          entry['OCPD'].indexOf("Force") !== -1 ? entry['OCPD'].slice(entry['OCPD'].indexOf("<= ") + 3, entry['OCPD'].indexOf(" cal/cm2")) : "N/A";
        const label_quantity = entry['Label Quantity'];
        if (!name || !distance_ft || !source || !location || !ocpdAmps || !ocpdType || !ocpdClass || !label_quantity) {
          throw new Error(`Missing required field(s) in entry: ${JSON.stringify(entry)}. Please ensure all required fields are filled.`);
        }
        ebmJSONData.push({
          name,
          location,
          distance_ft,
          source,
          ocpd: {
            amps: ocpdAmps,
            type: ocpdType,
            class: ocpdClass
          },
          label_quantity
        });
    } catch (error) {
      console.log('Error processing entry:', entry, error.message);
  }  
});
  return { end_use_equipment: ebmJSONData };
}