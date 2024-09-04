import { NextApiRequest, NextApiResponse } from 'next';
import multer from 'multer';
import fs from 'fs';
import { Document, Packer } from 'docx'; // Import Packer here
import { processCSVOutput } from '../processEstrellas';

// Set up multer to use a temporary directory
const upload = multer({ dest: '/tmp/' }); // Use '/tmp' for serverless compatibility

// Create a middleware function for multer
const multerMiddleware = upload.single('csvfile');

function runMiddleware(req: NextApiRequest, res: NextApiResponse, fn: Function) {
  return new Promise((resolve, reject) => {
    fn(req, res, (result: any) => {
      if (result instanceof Error) {
        return reject(result);
      }
      return resolve(result);
    });
  });
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method === 'POST') {
    try {
      await runMiddleware(req, res, multerMiddleware);

      const filePath = req.file?.path;

      if (!filePath) {
        return res.status(400).send('File not uploaded');
      }

      const estrellas: any[] = [];
      fs.createReadStream(filePath)
        .pipe(require('csv-parser')())
        .on('data', (row) => {
          estrellas.push(row);
        })
        .on('end', async () => {
          const { doc } = processCSVOutput(estrellas);

          // Return the DOCX as a response
          res.setHeader('Content-Disposition', 'attachment; filename=estrellas.docx');
          res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          );

          const buffer = await Packer.toBuffer(doc);
          res.send(buffer);
        });
    } catch (error) {
      res.status(500).send('Error processing file');
    }
  } else {
    res.status(405).send('Method Not Allowed');
  }
}
