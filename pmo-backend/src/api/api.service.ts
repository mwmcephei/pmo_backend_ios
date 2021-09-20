import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Sheet, SheetSchema } from '../schemas/sheet.schema';
import { Measure, MeasureSchema } from '../schemas/measure.schema';
import { Artefact, ArtefactSchema } from '../schemas/artefact.schema';
import { Budget, BudgetSchema } from '../schemas/budget.schema';
import { Model } from 'mongoose';
import { fileNames } from 'src/globalVars';
import '../types';
import { Overview } from '../types';

@Injectable()
export class ApiService {
  constructor(
    @InjectModel('Artefact') private artefactModel: Model<Artefact>,
    @InjectModel('Measure') private measureModel: Model<Measure>,
    @InjectModel('Sheet') private sheetModel: Model<Sheet>,
    @InjectModel('Budget') private budgetModel: Model<Budget>,
  ) { }

  async getMeasure(measureID: string): Promise<Measure> {
    try {
      const measure = await this.measureModel.findById(measureID);
      return measure;
    } catch (error) {
      return error;
    }
  }

  async getArtefactsOfMeasure(measureID: string): Promise<Artefact[]> {
    try {
      const measure = await this.measureModel
        .findById(measureID)
        .sort({ id: 'asc' });
      const populatedMeasure = await measure
        .populate('artefacts')
        .execPopulate();
      return populatedMeasure.artefacts;
    } catch (error) {
      return error;
    }
  }

  async getAllMeasures(): Promise<Measure[]> {
    try {
      const result = await this.measureModel.find().sort({ id: 'asc' });
      return result;
    } catch (error) {
      return error;
    }
  }

  async getOverview(): Promise<Sheet> {
    try {
      const excelSheet = await this.sheetModel.findOne({
        name: fileNames.main_file,
      });
      return excelSheet;
    } catch (error) {
      return error;
    }
  }

  async getBudget(): Promise<Budget> {
    try {
      const budget = await this.budgetModel.findOne();
      return budget;
    } catch (error) {
      return error;
    }
  }
}
