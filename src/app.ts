import express, { Request, Response } from 'express';
import path from 'path';
import cors from 'cors';
import { type DocLinkBody } from './types/types';
import { google } from 'googleapis';
import { extractDocId } from './utils/extractDocId';
import fs from 'fs';
import mammoth from 'mammoth';
import Excel, { type CellValue } from 'exceljs';
import { type SheetData } from './types/types';

const app = express();
app.use(express.json());
app.use(cors());

const port = 3001;

app.get('/', (req: Request, res: Response) => {
  res.send('Hello World from Express and TypeScript!');
});

const keyFilePath = path.join('docsparsing-12739e641ac7.json');
const auth = new google.auth.GoogleAuth({
  keyFile: keyFilePath,
  scopes: ['https://www.googleapis.com/auth/drive'],
});

const drive = google.drive({ version: 'v3', auth });

app.post('/sheets', async (req: Request, res: Response) => {
  const body: DocLinkBody = req.body;
  const { docLink } = body;
  console.log('Received google sheets doc link', docLink);

  const documentId = extractDocId(docLink);

  if (!documentId) {
    return res.status(400).json({ message: 'Invalid google doc link' });
  }

  try {
    const response = await drive.files.get(
      {
        fileId: documentId,
        alt: 'media',
      },
      { responseType: 'stream' }
    );

    const filePath = path.join(__dirname, 'temp.docx');
    const dest = fs.createWriteStream(filePath);

    response.data
      .on('end', () => {
        console.log('File downloaded');
        mammoth
          .extractRawText({ path: filePath })
          .then(result => {
            console.log(result.value);
            res.status(200).json({ content: result.value });
          })
          .catch(err => {
            console.error('Error reading docx.file', err);
            res.status(500).json({ message: 'Error reading the document' });
          })
          .finally(() => {
            fs.unlinkSync(filePath);
          });
      })
      .on('error', err => {
        console.error('Error downloading file', err);
        res.status(500).json({ message: 'Error downloading the document' });
      })
      .pipe(dest);
  } catch (error) {
    console.error('Failed to fetch the document', error);
    res.status(500).json({ message: 'Failed to fetch the document' });
  }
});

app.post('/tables', async (req: Request, res: Response) => {
  const body: DocLinkBody = req.body;
  const { docLink } = body;
  console.log('Received google sheets doc link', docLink);
  const documentId = extractDocId(docLink);

  if (!documentId) {
    return res.status(400).json({ message: 'Invalid google doc link' });
  }

  try {
    const response = await drive.files.get(
      {
        fileId: documentId,
        alt: 'media',
      },
      { responseType: 'stream' }
    );

    const filePath = path.join(__dirname, 'temp.xlsx');
    const dest = fs.createWriteStream(filePath);
    response.data
      .on('end', async () => {
        try {
          const workBook = new Excel.Workbook();
          await workBook.xlsx.readFile(filePath);
          const sheetsData: SheetData[] = [];

          workBook.eachSheet(sheet => {
            const sheetData: SheetData = { name: sheet.name, data: [] };
            sheet.eachRow({ includeEmpty: false }, row => {
              sheetData.data.push(row.values as CellValue[]);
            });
            sheetsData.push(sheetData);
          });
          res.status(200).json({ sheets: sheetsData });
        } catch (err) {
          console.error('Error reading xlsx file', err);
          res.status(500).json({ message: 'Failed to read xlsx file' });
        } finally {
          fs.unlinkSync(filePath);
        }
      })
      .on('error', err => {
        console.error('Error downloading the file:', err);
        res.status(500).json({ message: 'Failed to download the document' });
      })
      .pipe(dest);
  } catch (error) {
    console.error('Failed to fetch the document:', error);
    res.status(500).json({ message: 'Failed to fetch the document' });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
