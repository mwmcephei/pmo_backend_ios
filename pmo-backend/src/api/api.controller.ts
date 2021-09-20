import { Controller, Get, Param } from '@nestjs/common';
import { ApiService } from './api.service';

@Controller('api')
export class ApiController {
  constructor(private readonly apiService: ApiService) {}

  @Get('test')
  test() {
    return 'test';
  }

  @Get('measure/:measureID')
  getMeasure(@Param() params) {
    return this.apiService.getMeasure(params.measureID);
  }

  @Get('measure/:measureID/artefacts')
  getArtefactsOfMeasure(@Param() params) {
    return this.apiService.getArtefactsOfMeasure(params.measureID);
  }

  @Get('overview')
  getOverview() {
    return this.apiService.getOverview();
  }

  @Get('measures')
  getAllMeasures() {
    return this.apiService.getAllMeasures();
  }

  @Get('budget')
  getBudget() {
    return this.apiService.getBudget();
  }
}
