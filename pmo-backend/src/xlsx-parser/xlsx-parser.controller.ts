import { Controller, Get, Param } from '@nestjs/common';
import { XlsxParserService } from "./xlsx-parser.service";

@Controller('xlsx-parser')
export class XlsxParserController {
    constructor(private xlsxParseService: XlsxParserService) { }


    // routes for creating data
    //
    @Get('parse')
    parse() {
        console.log("parse")
        return this.xlsxParseService.parse();
    }

    @Get('parse_overview')
    parse_overview() {
        console.log("parse_overview")
        return this.xlsxParseService.parse_overview();
    }

    @Get("create_overview")
    triggerOverviewCreation() {
        console.log("createOverview")
        return this.xlsxParseService.createOverview();
    }


    @Get("parse_kpi")
    parseKPI() {
        console.log("parse_kpi")
        return this.xlsxParseService.parseKPI();
    }

    @Get("parse_budget_months")
    parseBudgetMonths() {
        console.log("parse_budget_months")
        return this.xlsxParseService.parseBudgetMonths();
    }


    // routes for retrieving data 
    //
    @Get('artefacts/:measureID')
    getArtefactsOfMeasure(@Param() params) {
        console.log("getAllArtefacts")
        console.log(params.measureID)
        return this.xlsxParseService.getArtefactsOfMeasure(params.measureID);
    }

    @Get("all_artefacts")
    getAllArtefacts() {
        console.log("getAllArtefacts")
        return "hi"
        //       return this.xlsxParseService.getAllArtefacts();
    }

    @Get("overview")
    getOverview() {
        console.log("getOverview")

        return this.xlsxParseService.getOverview();
    }

    @Get("measures")
    getAllMeasures() {
        console.log("getAllMeasures")
        return this.xlsxParseService.getAllMeasures();
    }

    @Get("budget")
    getBudget() {
        console.log("getBudget")
        return this.xlsxParseService.getBudget();
    }



}

