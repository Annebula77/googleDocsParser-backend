import { CellValue } from 'exceljs';

export interface DocLinkBody {
  docLink: string;
}

export interface SheetData {
  name: string;
  data: CellValue[][];
}
