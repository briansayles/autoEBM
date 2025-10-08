import { __dirname, __filename } from '../config.js';
import ebmStaticValues from '../static_data/ebmStaticValues.json' with {type: 'json'};
import shockBoundaries from '../static_data/shockBoundaries.json' with {type: 'json'};
import arcFlashBoundaries from '../static_data/arcFlashBoundaries.json' with {type: 'json'};
import XLSX from '@e965/xlsx';
import PizZip from 'pizzip';
import * as fs from 'fs';
XLSX.set_fs(fs);
import Docxtemplater from 'docxtemplater';
import expressionParser from "docxtemplater/expressions.js";
import DocxMerger from 'docx-merger';
import archiver from 'archiver';
import fsPromises from 'fs/promises';

export async function applyEnergyBoundaryMethod({dataFileName, noExcel, noLabels, noMerge, jobNumber, templateFileName} = {}) {
  const start0 = new Date().getTime();
  var start = new Date().getTime();
  const createExcel = noExcel == 'false'
  const createIndividualLabels = noLabels == 'false';
  const createMergeFile = noMerge == 'false';
  jobNumber = jobNumber || "";
  ebmStaticValues.sort((a, b)=> b.kA - a.kA);
  let dataFile;
  try {
    dataFile = await readEnergyBoundaryEntriesFromXLSX(dataFileName);
  } catch (error) {
    return {
        message: error.message,
        error: true,
        zipOutput: null,
    };
  }
  // console.log(`Data file read successfully `);
  // console.log(`Customer: ${dataFile.customer}`);
  // console.log(`Sources: ${dataFile.sources.map(source => source.name).join(', ')}`);
  // console.log(`AFIEs: ${dataFile.AFIEs.map(ie => ie.name).join(', ')}`);

  const templateData = await fsPromises.readFile(templateFileName, "binary");
  if (!templateData) {
    throw new Error(`Template file ${templateFileName} could not be read. Please check the template file.`);
  }
  const outputVariables = [];
  const excelOutputs = [];
  const nonEBMExcelOutputs = [];
  const nonEBMOutputVariables = [];
  const customerName = dataFile.customer;
  const sources = dataFile.sources;
  const ieArray = [];
  // console.log(`Processing autoEBM for customer: ${customerName}, job number: ${jobNumber}, with ${dataFile.end_use_equipment.length} EBM entries and ${dataFile.non_ebm_equipment.length} non-EBM entries.`);
  dataFile.AFIEs.forEach(ie_breakpoint => ieArray.push(ie_breakpoint.calories));
  // console.log(`Customer AFIEs (cal/cm2): ${ieArray.join(', ')}`);
  // console.log(dataFile.end_use_equipment.length > 0 ? `Processing ${dataFile.end_use_equipment.length} EBM entries.` : 'No EBM entries to process.');
  // console.log(dataFile.non_ebm_equipment.length > 0 ? `Processing ${dataFile.non_ebm_equipment.length} non-EBM entries.` : 'No non-EBM entries to process.');
  // console.log(dataFile);
  // Process each equipment item
  // return new Promise((resolve, reject) => {
  try {
    if (dataFile.end_use_equipment.length > 0) {
        dataFile.end_use_equipment.forEach((equipmentItem, equipmentIndex, equipmentArray) => {
          let recommendation = "";
          let recommendRK1 = false;
          let equipmentPPELevelRK1 = "";
          let source;
          let ebmStaticLine;
          let equipmentIEBreakpoint;
          let equipmentMaxIE;
          let equipmentPPELevel;
          let equipmentWorkingDistance;
          let ieBreakPoints;
          // console.log(`Processing EBM item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}`);
          if (equipmentItem.ocpd.type.indexOf("Force") == -1) {
            source = sources.filter((source)=> source.name == equipmentItem.source)[0];
          } else {
            // console.log('Creating new source for FORCE OCPD type');
            const sourceVoltage = parseInt(equipmentItem.ocpd.type.match(/@ (.*?)V/)[1]);
            // console.log(equipmentItem.ocpd.type, sourceVoltage);
            source = {
              name: equipmentItem.source,
              kA: 0,
              voltage: sourceVoltage
            }
          }
          if (!source) {
            throw new Error(`Error processing item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}. No source found matching name ${equipmentItem.source}. Please check the data file and customer data.`);
          }
          ebmStaticLine = ebmStaticValues.filter((ebmStaticValue) => 
            (ebmStaticValue.kA <= source.kA) * (JSON.stringify(ebmStaticValue.ocpd) == JSON.stringify(equipmentItem.ocpd)) * (ebmStaticValue.voltage == source.voltage) 
          )[0];
          if (!ebmStaticLine) {
            throw new Error(`Error processing item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem.name}. No EBM data found for source ${source.name} with ${source.kA} kA, ${source.voltage} V, and OCPD ${equipmentItem.ocpd.amps}A ${equipmentItem.ocpd.type} (${equipmentItem.ocpd.class}). Please check the data file and customer data.`);
          }
          equipmentWorkingDistance = ebmStaticLine.working_distance_in;
          ieBreakPoints = ebmStaticLine.boundaries.filter((boundary) =>  ieArray.includes(boundary.calories));
          equipmentIEBreakpoint = ieBreakPoints.filter(breakpoint => convertToNumber(breakpoint.distance_ft) >= convertToNumber(equipmentItem.distance_ft))[0]
          equipmentPPELevel = dataFile.AFIEs.filter(ie_breakpoint => ie_breakpoint.calories == equipmentIEBreakpoint.calories)[0].name;
          equipmentMaxIE = equipmentIEBreakpoint.calories;
          if (equipmentMaxIE > ieBreakPoints[0].calories) {
            if(equipmentItem.ocpd.class == "RK5") {
              recommendation = `Warning: Calculated max IE of ${equipmentMaxIE} cal/cm2 exceeds customer's minimum PPE level of ${ieBreakPoints[0].calories} cal/cm2.`; 
              const ebmStaticLineRK1 = ebmStaticValues.filter((ebmStaticValue) => 
                (ebmStaticValue.kA <= source.kA) * (JSON.stringify(ebmStaticValue.ocpd) === JSON.stringify({...equipmentItem.ocpd, class: "RK1"})) * (ebmStaticValue.voltage == source.voltage)
              )[0];
              const ieBreakPointsRK1 = ebmStaticLineRK1.boundaries.filter((boundary) =>  ieArray.includes(boundary.calories));
              const equipmentIEBreakpointRK1 = ieBreakPointsRK1.filter(breakpoint => convertToNumber(breakpoint.distance_ft) >= convertToNumber(equipmentItem.distance_ft))[0]
              const equipmentMaxIERK1 = equipmentIEBreakpointRK1.calories;
              equipmentPPELevelRK1 = dataFile.AFIEs.filter(ie_breakpoint => ie_breakpoint.calories == equipmentIEBreakpointRK1.calories)[0].name;
              if (equipmentMaxIERK1 < equipmentMaxIE) {
                recommendation += ` Using a class RK1 fuse will reduce the required PPE level from ${equipmentPPELevel} to ${dataFile.AFIEs.filter(ie_breakpoint => ie_breakpoint.calories == equipmentMaxIERK1)[0].name}.`;
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
          const today = new Date();
            outputVariables.push(
            {
              dataProvided: equipmentItem,
              timestamp: new Date().toISOString(),
              datestamp: `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`,
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
              ...equipmentItem, 
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
      }
      if (dataFile.non_ebm_equipment.length > 0) {
        dataFile.non_ebm_equipment.forEach((equipmentItem, equipmentIndex, equipmentArray) => {
          // console.log(`Processing non-EBM item ${equipmentIndex + 1} of ${equipmentArray.length}: ${equipmentItem['Title (Equipment Name)']}`);
          outputVariables.push(
            {
              ...equipmentItem,
              timestamp: new Date().toISOString(),
              datestamp: new Date().toISOString().slice(0,10),
              varAFB: equipmentItem["Arc Flash Boundary (in)"],
              varAFBFeetInches: equipmentItem["Arc Flash Boundary (ft-in)"] || toFeetInches(equipmentItem["Arc Flash Boundary (in)"]),
              varVoltage: equipmentItem["Voltage (kV)"]*1000 || equipmentItem["Voltage (V)"],
              varkV: equipmentItem["Voltage (kV)"] || equipmentItem["Voltage (V)"] * 1000,
              varRAB: equipmentItem["Restricted Approach Boundary (in)"],
              varRABFeetInches: equipmentItem["Restricted Approach Boundary (ft-in)"] || toFeetInches(equipmentItem["Restricted Approach Boundary (in)"]),
              varLAB: equipmentItem["Limited Approach Boundary (in)"],
              varLABFeetInches: equipmentItem["Limited Approach Boundary (ft-in)"] || toFeetInches(equipmentItem["Limited Approach Boundary (in)"]),
              varEquipmentName: equipmentItem["Title (Equipment Name)"],
              varFedFrom: equipmentItem["Source"] != "" ? equipmentItem["Source"] : equipmentItem["Source Protective Device"],
              varEquipmentLocation: equipmentItem["Equipment Location (Columns)"] || "",
              varPPE: equipmentItem["PPE Level (Site-Specific)"],
              varMaxIE: equipmentItem["Incident Energy (cal/cm2)"],
              varWorkingDistance: equipmentItem["Working Distance (in)"],
              varWorkingDistanceFeetInches: equipmentItem["Working Distance (ft-in)"] || toFeetInches(equipmentItem["Working Distance (in)"]),
              varQuantity: equipmentItem["Label Quantity"] || 1,
              varJobNumber: jobNumber
            });
          if (createExcel) {
            nonEBMExcelOutputs.push(
            {
              ...equipmentItem, 
              "Timestamp": new Date().toISOString(),
              "Job Number": jobNumber,
            });
          }
        },);
      }
    } catch (error) {
      return {
        message: error.message,
        error: true,
        zipOutput: null,
      };    
  }
  // console.log(outputVariables);
  var end = new Date().getTime();
  var time = end - start;
  start = new Date().getTime();
  let excelResult;
  let wordResult;
  // let zipResult;
  let zipOutput;
  const finishTimestamp = new Date().toISOString().replace(/:/g, '-').slice(0,19);
  const outputFilePath = await fsPromises.mkdir(`./output/${finishTimestamp}`, { recursive: true });
  await fsPromises.mkdir(`./output/${finishTimestamp}/individual labels`, { recursive: true });

  if (createExcel) {
    try {
      excelResult = await saveToExcel(excelOutputs, nonEBMExcelOutputs, customerName, jobNumber, finishTimestamp, outputFilePath);
    } catch (err) {
      console.log(err.message);
    }
  }
  if (createIndividualLabels) {
    try {
      wordResult = await generateMailMergeDOCX(outputVariables, customerName, createMergeFile, jobNumber, finishTimestamp, templateData, outputFilePath);
    } catch (err) {
      console.log(err.message);
    }
  }
  try {
    zipOutput = await createOutputZip(jobNumber, finishTimestamp, dataFileName, outputFilePath);  
  } catch (err) {
    console.log(err.message);
  }

  end = new Date().getTime();
  time = end - start0;
  return {
    message: `autoEBM processing complete for ${dataFile.end_use_equipment.length} entries in ${time/1000} seconds`, 
    error: false,
    zipOutput: zipOutput,
  };
}

async function saveToExcel(excelOutputs, nonEBMExcelOutputs, customerName, jobNumber, finishTimestamp, filePath) {
  try {
    // console.log(`Saving to Excel at ${filePath}`);
    const start = new Date().getTime();
    const EBMworksheet = XLSX.utils.json_to_sheet(excelOutputs);
    const nonEBMWorksheet = XLSX.utils.json_to_sheet(nonEBMExcelOutputs);
    const workbook = XLSX.utils.book_new();
    if (excelOutputs.length > 0) {
      XLSX.utils.book_append_sheet(workbook, EBMworksheet, 'EBM Results');
    }
    if (nonEBMExcelOutputs.length > 0) {
      XLSX.utils.book_append_sheet(workbook, nonEBMWorksheet, 'Non-EBM Results');
    }
    const excelFilename = toFilenameFriendlyFormat(`${customerName} AF Results ${jobNumber !== "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}`);
    XLSX.writeFile(workbook, `${filePath}/${excelFilename}.XLSX`);
    const end = new Date().getTime();
    const time = end - start;
    console.log(`Save to Excel took ${time / 1000} seconds for ${excelOutputs.length} items.`);
    return(filePath);
  } catch (error) {
    return (error.message);
  }
}

async function generateMailMergeDOCX(data, customerName, createMergeFile, jobNumber, finishTimestamp, templateFile, filePath) {
  try {
    var start = new Date().getTime();
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
          fsPromises.writeFile(`${filePath}/individual labels/${filename}.docx`, buffer, []);
          if (createMergeFile) {
            for (let i = 0; i < item.varQuantity; i++) {
              docxFiles.push(buffer);
            }
          }
        }
        catch (err) {
          console.error("Error: ", err.message);
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
      mergeFilename = toFilenameFriendlyFormat(`${customerName} AF Labels ${jobNumber != "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}`);
      docxMerger.save('nodebuffer', (data) => {
        fs.writeFileSync(`${filePath}/${mergeFilename}.docx`, data, (err) => {
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
        mergeFilePath: createMergeFile ? `${filePath}/${mergeFilename}.docx` : 'No merged file created'
      }
    )
  } catch (error) {
    console.log('Error during mail merge process:', error.message);
    return {
      message: `Error during mail merge process: ${error.message}`,
      mergeFilePath: null,
    };
  }
}

async function createOutputZip(jobNumber, finishTimestamp, dataFileName, filePath) {
  const zipPath = `./output/autoEBM Output ${jobNumber !== "" ? '(' + jobNumber + ') ' : ""}${finishTimestamp}.zip`;
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });
    output.on('close', () => {
      console.log(`Created output zip file at: ${zipPath} (${archive.pointer()} total bytes)`);
      resolve(zipPath);
      fsPromises.rmdir(filePath, { recursive: true }).then(() => {
        console.log(`Cleaned up output directory: ${filePath}`);
      }).catch((err) => {
        console.error(`Error removing output directory ${filePath}:`, err.message);
      }) 
    });
    archive.on('error', (err) => {
      reject(err);
    });
    archive.pipe(output);
    archive.directory(filePath, false);
    archive.file(dataFileName, { name: `Uploaded Data File ${dataFileName.split('/').pop()}` });
    archive.finalize();
  });
}

async function readEnergyBoundaryEntriesFromXLSX(excelFilename) {
    let ebmEntriesJSON, ebmEntrySheet, ebmEntryBook, customerDataSheet, sourcesDataSheet, AFIEsDataSheet, customer, sources, AFIEs, nonEBMEntrySheet, nonEBMEntriesJSON;
    try {
      console.log('reading ' + excelFilename);
      ebmEntryBook = XLSX.readFile(`${excelFilename}`);
    } catch (error) {
      console.log(error.message);
      return (error.message);
    };
    try { 
      ebmEntrySheet = ebmEntryBook.Sheets['Order Form (EBM)']; 
      ebmEntriesJSON = XLSX.utils.sheet_to_json(ebmEntrySheet);
      // console.log(`Found ${ebmEntriesJSON.length} EBM entries.`);
    } catch (error) {
      console.log(error.message);
      return (error.message);
    };
    try {
      nonEBMEntrySheet = ebmEntryBook.Sheets['Order Form (Non-EBM)'];
      nonEBMEntriesJSON = XLSX.utils.sheet_to_json(nonEBMEntrySheet);
      // console.log(`Found ${nonEBMEntriesJSON.length} non-EBM entries.`);
    } catch (error) {
      console.log(error.message);
    };
    try { 
      customerDataSheet = ebmEntryBook.Sheets['Customer']; 
      customer = XLSX.utils.sheet_to_json(customerDataSheet);
      customer = customer[0].customer;
      // console.log(`Customer name: ${customer}`);
    } catch (error) {
      return ('Error reading customer name: ' + error.message);
    };
    try { 
      sourcesDataSheet = ebmEntryBook.Sheets['Sources']; 
      sources = XLSX.utils.sheet_to_json(sourcesDataSheet);
      // console.log(`Found ${sources.length} sources.`);
    } catch (error) {
      return ('Error reading sources: ' + error.message);
    };
    try { 
      AFIEsDataSheet = ebmEntryBook.Sheets['AFIEs']; 
      AFIEs = XLSX.utils.sheet_to_json(AFIEsDataSheet);
      // console.log(`Found ${AFIEs.length} AFIEs.`);
    } catch (error) {
      return ('Error reading AFIEs: ' + error.message);
    };
    let ebmJSONData = [];
    if (!ebmEntriesJSON || ebmEntriesJSON.length == 0) {
      console.log('No EBM Entries in the datafile.');
      //return (new Error(`Data file ${excelFilename} is not valid or contains no data. Please check the data file.`));
    } else {
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
              entry['OCPD'].indexOf("Force") !== -1 ? entry['OCPD'] : "N/A";
            const ocpdClass = entry['OCPD'].indexOf("Class RK5") !== -1 ? "RK5" :
              entry['OCPD'].indexOf("Class RK1") !== -1 ? "RK1" : 
              entry['OCPD'].indexOf("Force") !== -1 ? entry['OCPD'].match(/<= (.*?) cal/)[1] : 
              entry['OCPD'].match(/A (.*?) MCCB/)[1]
            if (Array.isArray(ocpdClass)) {
              ocpdClass = ocpdClass[1];
            }
            const label_quantity = entry['Label Quantity'];
            if (!name || !distance_ft || !source || !location || !ocpdType || !ocpdClass || !label_quantity) {
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
          return (null);
        }  
      });
    }
    
    // console.log(JSON.stringify(ebmJSONData, null, 2));
    // console.log(JSON.stringify(sources, null, 2));
    // console.log(JSON.stringify(AFIEs, null, 2));
    // console.log(JSON.stringify(nonEBMEntriesJSON, null, 2));
    // console.log(`Customer: ${customer}`)
    return({ 
      end_use_equipment: ebmJSONData,
      customer: customer,
      sources: sources,
      AFIEs: AFIEs,
      non_ebm_equipment: nonEBMEntriesJSON,
    });
  // });
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