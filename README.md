# Web-Based Project for Energy Boundary Method Application

This project provides a web front-end interface that allows users to upload a data file and trigger the execution of the `applyEnergyBoundaryMethod` function from the backend. The application is built using Node.js and Express for the server-side, and HTML, CSS, and JavaScript for the client-side.

## Project Structure

```
web-front-end
├── public
│   ├── index.html          # HTML structure of the web front-end
│   ├── styles
│   │   └── style.css       # CSS styles for the web front-end
│   └── scripts
│       └── app.js          # Client-side JavaScript code
├── src
|   ├── static_data         # Proprietary data needed for EBM application
│   ├── server.js           # Entry point for the server-side application
│   └── utils
│       └── fileHandler.js   # Utility functions for handling file uploads
├── package.json            # npm configuration file
├── .gitignore              # Files and directories to be ignored by Git
└── README.md               # Documentation for the project
```

## Setup Instructions

1. **Clone the repository**:
   ```
   git clone git@github.com:briansayles/autoEBM.git
   cd autoEBM
   ```
2. **Create necessary files (proprietary data) in static_data folder**:
      arcFlashBoundaries.json
      ebmStatucValues.json
      shockBoundaries.json

3. **Install dependencies**:
   ```
   npm install
   ```

4. **Run the server**:
   ```
   node src/server.js OR npm start
   ```

5. **Open the application**:
   Navigate to `http://localhost:3000` in your web browser.

## Usage

- Use the file input buttons to select a data file, customer setup file and label template file (in the required formats).
- Enter a job number or name in the text field (optional).
- Use the checkboxes to disable Excel summary file creation, all Word label files creation, or only the merged Word file creation.
- Click the "Upload" button to send the file to the server.
- The server will process the file and execute the `applyEnergyBoundaryMethod` function.
- Upon completion, the web page will display links to the output files, which you can click to download to your local machine.

## License

This project is licensed under the MIT License. See the LICENSE file for details.