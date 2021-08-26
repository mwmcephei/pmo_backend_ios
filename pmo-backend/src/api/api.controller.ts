import { Controller, Get, Param } from '@nestjs/common';
import { ApiService } from './api.service';

@Controller('api')
export class ApiController {
  constructor(private readonly apiService: ApiService) { }


  @Get('test')
  test() {
    return "test"
  }






  @Get('measures/:measureID/artefacts')
  getArtefactsOfMeasure(@Param() params) {
    console.log('get artefacts of ' + params.measureID);
    console.log(params.measureID);
    return this.apiService.getArtefactsOfMeasure(params.measureID);
  }

  @Get('overview')
  getOverview() {
    console.log('getOverview');
    return this.apiService.getOverview();
  }

  @Get('measures')
  getAllMeasures() {
    console.log('getAllMeasures');
    return this.apiService.getAllMeasures();
  }

  @Get('budget')
  getBudget() {
    console.log('getBudget');
    return this.apiService.getBudget();
  }
}
