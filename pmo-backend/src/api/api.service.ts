import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Sheet, SheetSchema } from '../schemas/sheet.schema';
import { Measure, MeasureSchema } from '../schemas/measure.schema';
import { Artefact, ArtefactSchema } from '../schemas/artefact.schema';
import { Budget, BudgetSchema } from '../schemas/budget.schema';
import { Model } from 'mongoose';
import { fileNames } from 'src/globalVars';

@Injectable()
export class ApiService {
    constructor(
        @InjectModel('Artefact') private artefactModel: Model<Artefact>,
        @InjectModel('Measure') private measureModel: Model<Measure>,
        @InjectModel('Sheet') private sheetModel: Model<Sheet>,
        @InjectModel('Budget') private budgetModel: Model<Budget>,
    ) { }


    async getArtefactsOfMeasure(measureID: string): Promise<string> {
        const measure = await this.measureModel.findById(measureID);
        if (measure) {
            const populatedMeasure = await measure
                .populate('artefacts')
                .execPopulate();
            return JSON.stringify(populatedMeasure.artefacts);
        } else {
            return JSON.stringify('Error');
        }
    }

    async getAllMeasures(): Promise<string> {
        const result = await this.measureModel.find().sort({ id: "asc" })
        console.log(result);
        return JSON.stringify(result);
    }

    async getOverview(): Promise<string> {
        const excelSheet = await this.sheetModel.findOne({
            name: fileNames.main_file,
        }); // TO DO: safeguard for duplicates
        console.log(excelSheet);
        return JSON.stringify(excelSheet);
    }

    async getBudget(): Promise<string> {
        const excelSheet = await this.budgetModel.findOne(); // TO DO: safeguard for duplicates
        console.log(excelSheet);
        return JSON.stringify(excelSheet);
    }
}
