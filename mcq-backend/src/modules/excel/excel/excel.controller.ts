import { Controller, Post, Body, Get, Param } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller()
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Get('excel')
  getExcelData(): any[] {
    return this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet1').slice(1); // Remove header row
  }

  @Get('questions')
  getQuestions(): any[] {
    const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet1').slice(1);
    return excelData.map(row => ({ question: row[1], options: row.slice(2, 6) }));
  }

  @Get('questions/:id')
  getQuestionById(@Param('id') id: string): any {
    const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet1').slice(1);
    const question = excelData.find(row => row[0] === parseInt(id));
    if (!question) {
      return 'Question not found';
    }
    return {
      question: question[1],
      options: question.slice(2, 6)
    };
  }

  @Post('answer/:id')
  async submitAnswer(@Param('id') id: string, @Body() body: { answer: string, candidateName: string }): Promise<string> {
    const excelDataSheet1 = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet1');
    const excelDataSheet2 = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2');
    
    const question = excelDataSheet1.find(row => row[0] === parseInt(id));
    if (!question) {
      return 'Question not found';
    }
    
    const correctAnswer = question[6];
    const validationValue = 5;
    const candidateRow = excelDataSheet2.find(row => row[1] === body.candidateName);
    if (!candidateRow) {
      return 'Candidate not found';
    }
    
    const questionIndex = parseInt(id) + 2; // Q1 is in the 3rd column
    candidateRow[questionIndex] = body.answer === correctAnswer ? validationValue : 0;
    
    const totalColumnIndex = 23;
    candidateRow[totalColumnIndex] = candidateRow.slice(3, 23).reduce((sum, val) => sum + (val || 0), 0);
    
    await this.excelService.writeJSONToExcel('../mcq-backend/Book.xlsx', 'Sheet2', excelDataSheet2);
    return 'Answer submitted successfully!';
  }

  @Get('marks/:candidateName')
  async getTotalMarks(@Param('candidateName') candidateName: string): Promise<any> {
    const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2').slice(1); // Remove header row
    const candidateRow = excelData.find(row => row[1] === candidateName);
    if (!candidateRow) {
      return 'Candidate not found';
    }
    const totalMarks = candidateRow[23];
    return { candidateName, totalMarks };
  }

  @Get('candidateinfo')
  getCandidateInfo(): any[] {
    const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2').slice(1); // Remove header row
    return excelData.map(row => ({ id: row[0], candidateName: row[1], candidateRegisterNumber: row[2], totalMarks:row[23] }));
  }

  @Get('admin-candidates')
  getAdminCandidates(): any[] {
    const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2').slice(1); // Remove header row
    return excelData.map(row => ({ id: row[0], candidateName: row[1], candidateRegisterNumber: row[2] }));
  }

  @Get('admin-candidates/:id')
  getAdminCandidateById(@Param('id') id: string): any {
    const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2').slice(1); // Remove header row
    return excelData.find(candidate => candidate[0] === parseInt(id));
  }

  @Post('admin-candidates')
  createAdminCandidate(@Body() body: { candidateName: string, candidateRegisterNumber: string }): any {
    if (body.candidateName === 'admin' && body.candidateRegisterNumber === 'adminhr123') {
      const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2');
      const lastId = excelData[excelData.length - 1][0];
      const newId = lastId + 1;
      const newCandidate = [newId, body.candidateName, body.candidateRegisterNumber];
      excelData.push(newCandidate);
      this.excelService.writeJSONToExcel('../mcq-backend/Book.xlsx', 'Sheet2', excelData);
      return newCandidate;
    } else {
      return 'Unauthorized';
    }
  }

  @Post('admin-candidates/:id')
  updateAdminCandidate(@Param('id') id: string, @Body() body: { candidateName: string, candidateRegisterNumber: string }): any {
    if (body.candidateName === 'admin' && body.candidateRegisterNumber === 'adminhr123') {
      const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2');
      const candidateIndex = excelData.findIndex(candidate => candidate[0] === parseInt(id));
      if (candidateIndex !== -1) {
        excelData[candidateIndex][1] = body.candidateName;
        excelData[candidateIndex][2] = body.candidateRegisterNumber;
        this.excelService.writeJSONToExcel('../mcq-backend/Book.xlsx', 'Sheet2', excelData);
        return excelData[candidateIndex];
      } else {
        return 'Candidate not found';
      }
    } else {
      return 'Unauthorized';
    }
  }

  @Post('admin-candidates/:id/delete')
  deleteAdminCandidate(@Param('id') id: string, @Body() body: any): any {
    if (body.candidateName === 'admin' && body.candidateRegisterNumber === 'adminhr123') {
      const excelData = this.excelService.convertExcelToJSON('../mcq-backend/Book.xlsx', 'Sheet2');
      const candidateIndex = excelData.findIndex(candidate => candidate[0] === parseInt(id));
      if (candidateIndex !== -1) {
        excelData.splice(candidateIndex, 1);
        this.excelService.writeJSONToExcel('../mcq-backend/Book.xlsx', 'Sheet2', excelData);
        return { message: 'Candidate deleted successfully!' };
      } else {
        return 'Candidate not found';
      }
    } else {
      return 'Unauthorized';
    }
  }
}