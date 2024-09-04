const express = require('express');
const multer = require('multer');
const fs = require('fs');
const { processCSVOutput } = require('./processEstrellas');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.post('/upload', upload.single('csvfile'), (req, res) => {
  const filePath = req.file.path;

  // Read CSV file
  const estrellas = [];
  fs.createReadStream(filePath)
    .pipe(require('csv-parser')())
    .on('data', (row) => {
      estrellas.push(row);
    })
    .on('end', () => {
      // Process CSV data
      const { doc, text } = processCSVOutput(estrellas);

      // Send DOCX as response
      res.setHeader(
        'Content-Disposition',
        'attachment; filename=estrellas.docx'
      );
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      
      res.send(doc);
    });
});

app.listen(3000, () => {
  console.log('Server running on port 3000');
});

module.exports = app;
