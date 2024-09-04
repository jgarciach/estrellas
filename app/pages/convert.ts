import { NextApiRequest, NextApiResponse } from 'next';
import fs from 'fs';
import csv from 'csv-parser';
import { Document, Packer } from 'docx';
import { processCSVOutput } from '../../utils/csvProcessor';

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== 'POST') {
    return res.status(405).end();
  }

  const estrellas: any[] = [];

  req.pipe(csv())
    .on('data', (row) => {
      // Process each row and add to estrellas array
      // (Use the row processing logic from your original script)
    })
    .on('end', () => {
      const doc = processCSVOutput(estrellas);
      Packer.toBuffer(doc).then((buffer) => {
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename=estrellas.docx');
        res.send(buffer);
      });
    });
}