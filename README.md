# Web Front-End Project for Energy Boundary Method

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
   git clone <repository-url>
   cd web-front-end
   ```

2. **Install dependencies**:
   ```
   npm install
   ```

3. **Run the server**:
   ```
   node src/server.js
   ```

4. **Open the application**:
   Navigate to `http://localhost:3000` in your web browser.

## Usage

- Use the file input to select a data file (in the required format).
- Click the "Upload" button to send the file to the server.
- The server will process the file and execute the `applyEnergyBoundaryMethod` function.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.