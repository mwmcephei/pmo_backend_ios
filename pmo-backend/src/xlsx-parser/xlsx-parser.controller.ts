import { Controller, Get, Param } from '@nestjs/common';
import { XlsxParserService } from './xlsx-parser.service';

@Controller('xlsx-parser')
export class XlsxParserController {
  constructor(private xlsxParseService: XlsxParserService) {}

  // routes for creating data
  //
  @Get('parse')
  parse() {
    console.log('parse');
    return this.xlsxParseService.parse();
  }

  @Get('parse_overview')
  parse_overview() {
    console.log('parse_overview');
    return this.xlsxParseService.parse_overview();
  }

  @Get('create_overview')
  triggerOverviewCreation() {
    console.log('createOverview');
    return this.xlsxParseService.createOverview();
  }

  @Get('parse_kpi')
  parseKPI() {
    console.log('parse_kpi');
    return this.xlsxParseService.parseKPI();
  }

  @Get('parse_budget_months')
  parseBudgetMonths() {
    console.log('parse_budget_months');
    return this.xlsxParseService.parseBudgetMonths();
  }
}
