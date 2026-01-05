import express from 'express';
import fileUpload from 'express-fileupload';
import { applyEnergyBoundaryMethod } from './utils/autoEBM.js'; // Adjust the import based on your actual file structure
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs/promises';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Empties the upload directory on server start.
const uploadDir = path.join(__dirname, '../upload');
fs.readdir(uploadDir)
    .then(files => {
        return Promise.all(files.map(file => fs.unlink(path.join(uploadDir, file))));
    })
    .then(() => {
        console.log('Upload directory cleaned on server start.');
    })
    .catch(err => {
        console.error('Error cleaning upload directory:', err);
    }
);

// Empties the upload directory on server start.
const outputDir = path.join(__dirname, '../output');
fs.readdir(outputDir)
    .then(files => {
        return Promise.all(files.map(file => fs.unlink(path.join(outputDir, file))));
    })
    .then(() => {
        console.log('Output directory cleaned on server start.');
    })
    .catch(err => {
        console.error('Error cleaning output directory:', err);
    }
);

// Middleware to parse JSON bodies
app.use(express.json());


// Middleware for file uploads
app.use(fileUpload());

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, '../public')));

// Endpoint to handle file upload and process it
app.post('/upload', async (req, res) => {
    if (!req.files || Object.keys(req.files).length < 1) {
        return res.send({ message: 'Missing required files in upload.' });
    }
    console.log('File upload(s) received');
    const dataFile = req.files.equipmentDataFile;
    // const configFile = req.files.customerConfigFile;
    if (!dataFile) {
        return res.send({ message: 'Please upload all required files: equipmentDataFile' });
    }
    // console.log(`Received file: ${dataFile.name}`);
    // console.log(`Received config file: ${configFile.name}`);
    // console.log(`Received template file: ${templateFile.name}`);
    // Save the uploaded file temporarily
    const uploadPath = path.join(__dirname, '../upload', dataFile.name);
    await dataFile.mv(uploadPath);
    // const configUploadPath = path.join(__dirname, '../upload', configFile.name);
    // await configFile.mv(configUploadPath);
    if (!req.headers.noLabels) {
        const templateFile = req.files.customerTemplateFile;
        const templateUploadPath = path.join(__dirname, '../upload', templateFile.name);
        await templateFile.mv(templateUploadPath);
    } else {
        var templateUploadPath = "";
    }
    try {
        // Call the applyEnergyBoundaryMethod function with the uploaded file and other parameters from headers
        console.log('Calling applyEnergyBoundaryMethod')
        const autoEBMResult = await applyEnergyBoundaryMethod({ 
            dataFileName: uploadPath, 
            jobNumber: req.headers.jobnumber, 
            noExcel: req.headers.noexcel, 
            noLabels: req.headers.nolabels, 
            noMerge: req.headers.nomerge,
            // customerName: req.headers.customername,
            // configFileName: configUploadPath,
            templateFileName: templateUploadPath
        });
        res.send(autoEBMResult);
    } catch (error) {
        res.status(500).send({ message: error.message });
    } finally {
        console.log('Processing complete, cleaning up uploaded files if not already deleted.');
        // Delete the uploaded files after processing
        await fs.unlink(uploadPath).catch((unlinkError) => {
            console.error('Error deleting file:', unlinkError.message);
        });
        // await fs.unlink(configUploadPath).catch((unlinkError) => {
        //     console.error('Error deleting config file:', unlinkError.message);
        // });
        await fs.unlink(templateUploadPath).catch((unlinkError) => {
            console.error('Error deleting template file:', unlinkError.message);
        });
        console.log('Uploaded files deleted successfully.');
    }
    console.log('Upload endpoint processing complete');
    console.log('------------------------------------');
});

app.get('/download', async (req, res, next) => {
  const filePath = req.query.filePath; // File path passed as a query parameter

  if (!filePath) {
      return res.status(400).send('File path is required.');
  }

  res.download(filePath, (err) => {
      if (err) {
          console.error('Error downloading file:', err.message);
          return res.status(500).send('Error downloading file.');
      } else {
            console.log('File downloaded successfully:', filePath);
            // Delete the file after download
            fs.unlink(filePath).catch((unlinkError) => {
                console.error('Error deleting downloaded file:', unlinkError.message);
            });
      }
  });
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});