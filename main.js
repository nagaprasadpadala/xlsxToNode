const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const cors = require('cors');

const app = express();
const port = 3002;

app.use(cors());
// Set up Multer to handle file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Serve static files from the public folder
app.use(express.static('public'));

// Handle file upload
app.post('/upload', upload.single('file'), (req, res) => {
  // Access the uploaded file
  const fileBuffer = req.file.buffer;
  // console.log('File buffer:', fileBuffer);

  // Parse the XLSX file
  const workbook = xlsx.read(fileBuffer, { type: 'buffer' });

  // Process the workbook
  const sheetName = workbook.SheetNames[0]; // Assuming there is only one sheet
  const worksheet = workbook.Sheets[sheetName];

  // Convert worksheet to JSON
  
  const jsonData = xlsx.utils.sheet_to_json(worksheet);
  var multipliedData = [];
  for(let i=0; i<jsonData.length; i++){
    // console.log('test');
    // console.log(jsonData[i]['value1']);   
    // multipliedData.push(jsonData[i]);
    multipliedData.push({value1: jsonData[i]['value1'], value2: jsonData[i]['value2'], Product:(jsonData[i]['value1'] *  jsonData[i]['value2']), Addition:(jsonData[i]['value1'] +  jsonData[i]['value2'])});
  }
console.log(multipliedData);
  // Create a new XLSX file with '1', '2', and 'Result' columns
  const newWorkbook = xlsx.utils.book_new();
  const newWorksheet = xlsx.utils.json_to_sheet(multipliedData);
  xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet 1');

  // Save the new XLSX file to a buffer
  const newFileBuffer = xlsx.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });

  // Create a unique filename
  const fileName = `result.xlsx`;

  // Save the buffer to a file on the server
  fs.writeFileSync(fileName, newFileBuffer);

  // Send a JSON response with the filename to the client
  res.json({ fileName });
});

app.get('/download/:fileName', (req, res) => {
  const fileName = req.params.fileName;
  const filePath = `./${fileName}`;

  // Send the file as a response for download
  res.download(filePath, (err) => {
    // Cleanup: Delete the temporary file after it's sent
    fs.unlinkSync(filePath);
  });
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
