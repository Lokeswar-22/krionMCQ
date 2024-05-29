import * as XLSX from 'xlsx';
import { Injectable } from '@nestjs/common';

@Injectable()
export class ExcelService {
  convertExcelToJSON(filePath: string, sheetName: string): any[] {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    return data;
  }

  writeJSONToExcel(filePath: string, sheetName: string, jsonData: any[][]): void {
    const workbook = XLSX.readFile(filePath);
    const newWorksheet = XLSX.utils.aoa_to_sheet(jsonData);
    workbook.Sheets[sheetName] = newWorksheet;
    XLSX.writeFile(workbook, filePath);
  }
}
