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
  let ebmEntries;
  try {
    ebmEntries = await readEnergyBoundaryEntriesFromXLSX(dataFileName);
  } catch (error) {
    return {
      message: error.message,
      error: true,
      excelFilePath: null,
      wordFilePath: null,
      labelsZipPath: null
    };
  }
  const customerDataJSON = await import(configFileName, {with: {type: 'json'}});
  if (!customerDataJSON || !customerDataJSON.default || customerDataJSON.default.length == 0) {
    throw new Error(`Configuration file ${configFileName} is not valid. Please check the customer configuration.`);
  }
  const templateData = fs.readFileSync(templateFileName, "binary");
  const outputVariables = [];
  const excelOutputs = [];
  const customerData = customerDataJSON.default[0].customer;
  const sources = customerData.sources;
  const ieArray = [];
  customerData.ie_breakpoints.forEach(ie_breakpoint => ieArray.push(ie_breakpoint.calories));
  try {
      ebmEntries.end_use_equipment.forEach((equipmentItem, equipmentIndex, equipmentArray) => {
        let recommendation = "";
        let recommendRK1 = false;
        let equipmentPPELevelRK1 = "";
        const source = sources.filter((source)=> source.name == equipmentItem.source)[0];
        if (!source) {
          throw new Error(`Error processing item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}. No source found matching name ${equipmentItem.source}. Please check the data file and customer data.`);
        }
        const ebmStaticLine = ebmStaticValues.filter((ebmStaticValue) => 
          (ebmStaticValue.kA <= source.kA) * (JSON.stringify(ebmStaticValue.ocpd) === JSON.stringify(equipmentItem.ocpd)) * (ebmStaticValue.voltage == source.voltage) 
        )[0];
        if (!ebmStaticLine) {
          throw new Error(`Error processing item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}. No EBM data found for source ${source.name} with ${source.kA} kA, ${source.voltage} V, and OCPD ${equipmentItem.ocpd.amps}A ${equipmentItem.ocpd.type} (${equipmentItem.ocpd.class}). Please check the data file and customer data.`);
        }
        const equipmentWorkingDistance = ebmStaticLine.working_distance_in;
        const ieBreakPoints = ebmStaticLine.boundaries.filter((boundary) =>  ieArray.includes(boundary.calories));
        const equipmentIEBreakpoint = ieBreakPoints.filter(breakpoint => convertToNumber(breakpoint.distance_ft) >= convertToNumber(equipmentItem.distance_ft))[0]
        const equipmentPPELevel = customerData.ie_breakpoints.filter(ie_breakpoint => ie_breakpoint.calories == equipmentIEBreakpoint.calories)[0].name;
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

async function saveToExcel(excelOutputs, customer, jobNumber, finishTimestamp) {
  return new Promise((resolve, reject) => {
    try {
      const start = new Date().getTime();
      const worksheet = XLSX.utils.json_to_sheet(excelOutputs);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'AF Results');
      fs.mkdirSync(`./output/${finishTimestamp}`, { recursive: true });
      const excelFilename = toFilenameFriendlyFormat(`${customer.name} AF Results ${jobNumber !== "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}`);
      const filePath = `./output/${finishTimestamp}/${excelFilename}.xlsx`;
      XLSX.writeFile(workbook, filePath);
      const end = new Date().getTime();
      const time = end - start;
      console.log(`Save to Excel took ${time / 1000} seconds for ${excelOutputs.length} items.`);
      resolve(filePath);
    } catch (error) {
      reject(new Error(`Error saving to Excel: ${error.message}`));
    }
  });
}

async function generateMailMergeDOCX(data, customer, createMergeFile, jobNumber, finishTimestamp, templateFile) {
  return new Promise((resolve, reject) => {
    try {
      var start = new Date().getTime();
      fs.mkdirSync(`./output/${finishTimestamp}/individual labels/`, { recursive: true });
      var docxFiles = [];
      data.forEach((item, index, array) => {
        if (item.varQuantity > 0) {
          try {
            var filename = toFilenameFriendlyFormat(`${item.varEquipmentName} ${jobNumber != "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}`);
            item.varLabelID = `${item.varEquipmentName}-${finishTimestamp}`;
            const zip = new PizZip(templateFile);
            var doc = new Docxtemplater(zip, {
              parser: expressionParser,
              linebreaks: true,
              paragraphLoop: true,
            });
            doc.render(item);
            var buffer = doc.toBuffer();
            fs.writeFileSync(`./output/${finishTimestamp}/individual labels/${filename}.docx`, buffer, []);
            if (createMergeFile) {
              for (let i = 0; i < item.varQuantity; i++) {
                docxFiles.push(buffer);
              }
            }
          }
          catch (err) {
            console.error("Error: ", err.message);
            reject(new Error(`Error creating label file for ${item.varEquipmentName}: ${err.message}`));
            return;
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
      resolve (
        {
          message: 'Mail merge complete',
          mergeFilePath: createMergeFile ? `./output/${finishTimestamp}/${mergeFilename}.docx` : 'No merged file created'
        }
      )
    } catch (error) {
      console.log('Error during mail merge process:', error.message);
      reject(new Error(`Error during mail merge process: ${error.message}`));
    }
  });
}
  
async function createLabelsZip(jobNumber, finishTimestamp) {
  return new Promise((resolve, reject) => {
    try {
      const labelsDir = `./output/${finishTimestamp}/individual labels`;
      const zipPath = `./output/${finishTimestamp}/individual_labels ${jobNumber !== "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}.zip`;
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
    } catch (error) {
      reject(new Error(`Error creating labels zip: ${error.message}`));
    }
  });
}

async function readEnergyBoundaryEntriesFromXLSX(excelFilename) {
  return new Promise((resolve, reject) => {
    let ebmEntriesJSON, ebmEntrySheet, ebmEntryBook;
    try {
      ebmEntryBook = XLSX.readFile(`${excelFilename}`);
    } catch (error) {
      reject(new Error(`Error reading Excel file ${excelFilename}: ${error.message}. Please ensure the file is a valid .xlsx file.`));
      return;
    };
    try { 
      ebmEntrySheet = ebmEntryBook.Sheets['Order Form']; 
    } catch (error) {
      reject(new Error(`Error accessing 'Order Form' sheet in Excel file ${excelFilename}: ${error.message}. Please ensure the sheet exists and is named correctly.`));
      return;
    };
    try {
      ebmEntriesJSON = XLSX.utils.sheet_to_json(ebmEntrySheet);
    } catch (error) {
      reject(new Error(`Error converting 'Order Form' sheet to JSON in Excel file ${excelFilename}: ${error.message}. Please ensure the sheet is properly formatted.`));
      return;
    };
    if (!ebmEntriesJSON || ebmEntriesJSON.length == 0) {
      reject(new Error(`No data found in 'Order Form' sheet of Excel file ${excelFilename}. Please ensure the sheet contains data.`));
      return;
    }
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
        reject(new Error(`Error processing entry: ${JSON.stringify(entry)}. ${error.message}`));
        return;
      }  
    });
    resolve({ end_use_equipment: ebmJSONData });
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