import fs from 'fs';
import path from 'path';
// import { applyEnergyBoundaryMethod } from '../../../autoEBM/autoEBM.js';

export const handleFileUpload = (file) => {
  const filePath = path.join(__dirname, '../upload', file.name);
  return new Promise((resolve, reject) => {
    const writeStream = fs.createWriteStream(filePath);
    file.stream.pipe(writeStream);
    writeStream.on('finish', () => resolve(filePath));
    writeStream.on('error', (error) => reject(error));
  });
};

// export const processUploadedFile = async (filePath) => {
//   try {
//     const result = await applyEnergyBoundaryMethod(filePath);
//     return result;
//   } catch (error) {
//     throw new Error(`Error processing file: ${error.message}`);
//   }
// };

export const fileDownloadMiddleware = (req, res, next) => {
  const filePath = req.query.filePath; // File path passed as a query parameter

  if (!filePath) {
      return res.status(400).send('File path is required.');
  }

  if (!fs.existsSync(filePath)) {
      return res.status(404).send('File not found.');
  }

  res.download(filePath, (err) => {
      if (err) {
          console.error('Error downloading file:', err.message);
          return res.status(500).send('Error downloading file.');
      }
  });
};