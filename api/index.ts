const multer = require('multer');
const fs = require('fs');
const { processCSVOutput } = require('../processEstrellas');

const upload = multer({ dest: '/tmp/' }); // Use '/tmp' for serverless compatibility

// Express-like request handler for Vercel
export default async (req, res) => {
  if (req.method === 'POST') {
    upload.single('csvfile')(req, res, (err) => {
      if (err) {
        return res.status(500).send('File upload error');
      }

      const filePath = req.file.path;
      const estrellas = [];

      fs.createReadStream(filePath)
        .pipe(require('csv-parser')())
        .on('data', (row) => {
          estrellas.push(row);
        })
        .on('end', () => {
          const { doc } = processCSVOutput(estrellas);

          // Return the DOCX as a response
          res.setHeader(
            'Content-Disposition',
            'attachment; filename=estrellas.docx'
          );
          res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          );

          Packer.toBuffer(doc).then((buffer) => {
            res.send(buffer);
          });
        });
    });
  } else {
    res.status(405).send('Method Not Allowed');
  }
};
