import { Injectable } from '@nestjs/common';
import * as fs from 'fs';
import { InjectModel } from '@nestjs/mongoose';
import { Sheet, SheetSchema } from '../schemas/sheet.schema';
import { Measure, MeasureSchema } from '../schemas/measure.schema';
import { Artefact, ArtefactSchema } from '../schemas/artefact.schema';
import { Budget, BudgetSchema } from '../schemas/budget.schema';
import { Model } from 'mongoose';
import { resolve } from 'path';
import * as XLSX from 'xlsx';
import { fileNames } from '../globalVars';


/*
Conduct one-time manual parsing by addressing api endpoints in this order:
1. .../xlsx-parser/parse
2. .../xlsx-parser/parse_overview
3. .../xlsx-parser/create_overview  
4. .../xlsx-parser/parse_kpi
5. .../xlsx-parser/parse_budget_months
*/

type Overview = {
  numberOfMeasures: number,
  totalBudget: number,
  overallStatus: number,
  progressOverviewBarResult: number,
  KPIProgressResult: number,
}
type ParseOverview = {
  row: number,
  name: string,
  risk: number,
  budget: number,
  artefact: number
}
type KpiProgressData = {
  measureName: string,
  progress: number
}
type AllBudgetMeasures = { [x: number]: number }





type SheetType = {
  [key: string]: string | number
}

@Injectable()
export class XlsxParserService {
  constructor(
    @InjectModel('Artefact') private artefactModel: Model<Artefact>,
    @InjectModel('Measure') private measureModel: Model<Measure>,
    @InjectModel('Sheet') private sheetModel: Model<Sheet>,
    @InjectModel('Budget') private budgetModel: Model<Budget>,
  ) { }



  async createOverview(): Promise<Overview> {
    let result: Overview;
    const excelSheet = await this.sheetModel.findOne({
      name: fileNames.main_file,
    });
    if (excelSheet) {
      const numberOfMeasures = excelSheet.measures.length;
      const workbook = XLSX.readFile(
        resolve(fileNames.xlsx_file_dir, fileNames.main_file),
      );
      const overview_object = workbook.Sheets['Status Overview'];

      // ------total budget
      let totalBudget = 0;
      const allBudgetsOfMeasures: AllBudgetMeasures[] = [];
      Object.keys(overview_object).filter((key) => {
        if (key.includes('I')) {      // column 'I' of xlsx sheet
          const row = parseInt(key.substring(1));
          if (row > 4) {
            const measureName = overview_object['D' + row]['v'];
            const budgetAsString = overview_object[key]['v'];
            const budget =
              parseInt(
                budgetAsString.substring(0, budgetAsString.indexOf('k')),
              ) * 1000;
            totalBudget += budget;
            const currentBudget = {
              [measureName]: budget,
            };
            allBudgetsOfMeasures.push(currentBudget);
          }
        }
      });
      // ------overall status
      const allRisksBudgetsAndArtefacts = [];
      Object.keys(overview_object).filter((key) => {
        if (key.includes('P')) {
          const row = parseInt(key.substring(1));
          if (row > 4) {
            allRisksBudgetsAndArtefacts.push(overview_object[key]['v']);
          }
        }
      });
      Object.keys(overview_object).filter((key) => {
        if (key.includes('Q')) {
          const row = parseInt(key.substring(1));
          if (row > 4) {
            allRisksBudgetsAndArtefacts.push(overview_object[key]['v']);
          }
        }
      });
      Object.keys(overview_object).filter((key) => {
        if (key.includes('R')) {
          const row = parseInt(key.substring(1));
          if (row > 4) {
            allRisksBudgetsAndArtefacts.push(overview_object[key]['v']);
          }
        }
      });
      let overallStatus = 0;
      allRisksBudgetsAndArtefacts.map((a) => {
        if (a > overallStatus) {
          overallStatus = a;
        }
      });

      // ------Progress Overview:  sum over measures(avgProgress * measureBudget) / totalBudget
      // get all measures, get artefacts of each measure 
      let sumAvgProgressTimesBudgetOfMEasures = 0;
      for (let m = 0; m < excelSheet.measures.length; m++) {
        const measure = await (
          await this.measureModel.findById(excelSheet.measures[m])
        )
          .populate('artefacts')
          .execPopulate();
        let avgProgressOfArtefacts = 0;
        measure.artefacts.map((art) => {
          avgProgressOfArtefacts += art.progress;
        });
        avgProgressOfArtefacts =
          avgProgressOfArtefacts / measure.artefacts.length;
        let temp = 0;
        allBudgetsOfMeasures.map((item) => {
          if (item[measure.title]) {
            temp = item[measure.title] * avgProgressOfArtefacts;
          }
        });
        sumAvgProgressTimesBudgetOfMEasures += temp;
      }
      const progressOverviewBarResult =
        sumAvgProgressTimesBudgetOfMEasures / totalBudget;

      // ------KPI Progress
      const KPIprogressOfAllMeasures =
        this.getKPIProgressData('kpi_progress.xlsx');
      let sum = 0;
      KPIprogressOfAllMeasures.map((item) => {
        allBudgetsOfMeasures.map((budgetOfMeasure) => {
          if (budgetOfMeasure[item.measureName]) {
            const temp = item.progress * budgetOfMeasure[item.measureName];
            sum += temp;
          }
        });
      });
      const KPIProgressResult = sum / totalBudget;
      const updatedSheet = await excelSheet.update({
        totalBudget: totalBudget,
        overallStatus: overallStatus,
        progress: Math.round(progressOverviewBarResult * 100) / 100,
        kpiProgress: Math.round(KPIProgressResult * 100) / 100,
      })
      if (updatedSheet) {
        console.log('updated')
        console.log(updatedSheet)
      }
      result = {
        numberOfMeasures,
        totalBudget,
        overallStatus,
        progressOverviewBarResult,
        KPIProgressResult,
      }
    }
    return result;
  }

  // aux function for createOverview()
  getKPIProgressData(kpiFile: string): KpiProgressData[] {
    const workbook = XLSX.readFile(resolve(fileNames.xlsx_file_dir, kpiFile));
    const overview_object = workbook.Sheets['Plan view'];
    // D:measure name, G current progress, H target progress
    const numberOfRows = 22;        // TO DO: get number of rows programmatically    ---   22
    const result: KpiProgressData[] = [];
    for (let i = 1; i <= numberOfRows; i++) {
      const keyMeasureName = 'D' + (4 + i);          // first entry at row 4
      const keyActualProgress = 'G' + (4 + i);
      const keyTargetProgress = 'H' + (4 + i);
      result.push({
        measureName: overview_object[keyMeasureName]['v'],
        progress:
          Math.round(
            (overview_object[keyActualProgress]['v'] /
              overview_object[keyTargetProgress]['v']) *
            100,
          ) / 100,
      });
    }
    return result;
  }


  // parse and save measures and corresponding artefacts
  parse(): string {
    // create Sheet Table
    const newSheet = {
      name: fileNames.main_file,
    }
    const excelFile = new this.sheetModel(newSheet)
    excelFile.save()
      .then(newlySavedExcelSheet => {
        console.log(newlySavedExcelSheet)

        const workbook = XLSX.readFile(resolve(fileNames.xlsx_file_dir, fileNames.main_file))
        // 'sheet' corresponds to measure
        const sheet_name_list = workbook.SheetNames;
        console.log(typeof (workbook.SheetNames))
        sheet_name_list.map(sheetName => {
          // save measure to DB
          if (sheetName !== "Status Overview" && sheetName !== "Overview") {
            console.log(sheetName)
            const newMeasure = {
              title: sheetName,
            }
            const measure = new this.measureModel(newMeasure)
            measure.save()
              .then(async savedMeasure => {
                // add measure to ExcelSheet measure list
                await this.sheetModel.updateOne({ _id: newlySavedExcelSheet._id }, { $push: { measures: savedMeasure } });
                return savedMeasure
              })
              .then(async savedMeasure => {
                // get artefacts of this measure and add it to measure in DB
                const data: SheetType[] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName])
                const artefacts = this.getArtefactsFromLinesArray(data)

                const savedArtefact_IDs = []
                artefacts.map(art => {  // artefacts: array of objects, each containing a row of xlsx file
                  const toSave = {
                    id: art["__EMPTY_1"], // __EMPTY_ + column(!) number accesses a cell
                    description: art["__EMPTY_2"],
                    progress: art["__EMPTY_9"],
                    budget: art["__EMPTY_11"] ? art["__EMPTY_11"] : "",
                    achievement: art["__EMPTY_13"],
                    work: art["__EMPTY_21"],
                  }
                  const artefact = new this.artefactModel(toSave)
                  artefact.save()
                    .then(async savedArtefact => {
                      savedArtefact_IDs.push(savedArtefact._id)
                      await this.measureModel.updateOne({ _id: savedMeasure._id }, { $push: { artefacts: savedArtefact } });
                    })
                    .catch(err => console.log(err))
                })
              })
              .catch(err => console.log(err))
          }
        })
      })
    return "measures & artefacts parsed and saved to DB"
  }



  // aux function for parse()
  getArtefactsFromLinesArray(sheet: SheetType[]): SheetType[] {
    return sheet.filter(line => {
      const firstKey = Object.keys(line)[0]
      if (firstKey === "__EMPTY_1") {
        const firstItem = `${line[firstKey]}`
        try {
          if (parseInt(firstItem) < 10 && Object.keys(line).length > 2) {
            return line;
          }
        } catch (e) {
          console.log(e);
        }
      }
    });
  }



  // adds status info to measures
  parse_overview(): ParseOverview[] {
    const workbook = XLSX.readFile(
      resolve(fileNames.xlsx_file_dir, fileNames.main_file),
    );
    // parse overview
    const overview_object = workbook.Sheets['Status Overview'];
    const risks = []; // RESTRUCTURE!!!!!!!
    Object.keys(overview_object).filter((key) => {
      if (key.includes('P')) {
        const row = parseInt(key.substring(1));
        if (row > 4) {
          risks.push({
            row,
            risk: overview_object[key]['v'],
          });
        }
      }
    });
    const budgets = [];
    Object.keys(overview_object).filter((key) => {
      if (key.includes('Q')) {
        const row = parseInt(key.substring(1));
        if (row > 4) {
          budgets.push({
            row,
            budget: overview_object[key]['v'],
          });
        }
      }
    });
    const artefacts = [];
    Object.keys(overview_object).filter((key) => {
      if (key.includes('R')) {
        const row = parseInt(key.substring(1));
        if (row > 4) {
          artefacts.push({
            row,
            artefact: overview_object[key]['v'],
          });
        }
      }
    });
    const result = [];
    Object.keys(overview_object).filter(async (key) => {
      if (key.includes('D')) {
        const row = parseInt(key.substring(1));
        if (row > 4) {
          const addToResult = {
            row,
            name: overview_object[key]['h'],
            risk: risks[row - 5]['risk'],
            budget: budgets[row - 5]['budget'],
            artefact: artefacts[row - 5]['artefact'],
          };
          result.push(addToResult);
          await this.measureModel.updateOne(
            { title: addToResult.name },
            {
              risk: addToResult.risk,
              budget: addToResult.budget,
              artefact: addToResult.artefact,
            },
          );
        }
      }
    });
    return result;
  }


  async parseKPI(): Promise<string> {
    const workbook = XLSX.readFile(
      resolve(fileNames.xlsx_file_dir, fileNames.kpi_file_2),
    );
    const overview_object = workbook.Sheets['Plan view'];
    // get row number of measures. measure names in column "D"
    const rowsOfMeasures = [];
    Object.keys(overview_object).filter((key) => {
      if (key.includes('D')) {
        const row = parseInt(key.substring(1));
        if (row > 4) {
          rowsOfMeasures.push({
            measureName: overview_object[key]['v'],
            row,
          });
        }
      }
    });
    // get measures from DB
    const measures = await this.measureModel.find();
    measures.map(async (measure) => {
      let rowOfThisMeasure;
      for (let i = 0; i < rowsOfMeasures.length; i++) {
        if (rowsOfMeasures[i].measureName === measure.title) {
          rowOfThisMeasure = rowsOfMeasures[i].row;
        }
      }
      // from row get actual, target, plan of last month
      // G, H and J ???
      let actuals;
      let target;
      let lastPlan;
      Object.keys(overview_object).filter((key) => {
        if (key.includes('G')) {
          const row = parseInt(key.substring(1));
          if (row == rowOfThisMeasure) {
            actuals = overview_object[key]['v'];
          }
        }
        if (key.includes('H')) {
          const row = parseInt(key.substring(1));
          if (row == rowOfThisMeasure) {
            target = overview_object[key]['v'];
          }
        }
        if (key.includes('L')) {
          // TO DO: clarify which last month
          const row = parseInt(key.substring(1));
          if (row == rowOfThisMeasure) {
            lastPlan = overview_object[key]['v'];
          }
        }
      });
      console.log(
        measure.title + '  ' + actuals + '  ' + target + '  ' + lastPlan,
      );
      let kpiProgressOfThisMeasure;
      if (actuals < lastPlan) {
        kpiProgressOfThisMeasure = 0; // behind schedule
      } else if (lastPlan <= actuals && actuals < target) {
        kpiProgressOfThisMeasure = 1; // on schedule
      } else if (actuals >= target) {
        kpiProgressOfThisMeasure = 2; // finished
      }
      const updatedMeasure = await measure.update({
        kpiProgress: kpiProgressOfThisMeasure,
      })
      if (updatedMeasure) {
        console.log('updated');
        console.log(updatedMeasure)
      }
    });
    return 'ok';
  }

  async parseBudgetMonths(): Promise<string> {
    const workbook = XLSX.readFile(
      resolve(fileNames.xlsx_file_dir, fileNames.budget_file),
    );
    const overview_object = workbook.Sheets['1. Overview'];
    const detailes_object = workbook.Sheets['2. Detailed view'];
    // M > 5,  D measure names
    // M 29 grand total approved budget
    const totalApprovedBudget = overview_object['M29']['v'];
    // get month columns EUR1 EUR2 .... row 12
    const month_columns = [];
    Object.keys(detailes_object).map((key) => {
      const tmp = key.replace(/^[A-Z]/, '_');
      const split = tmp.split('_');
      const target = split[split.length - 1];
      if (parseInt(target) == 12) {
        const x = detailes_object[key]['v'];
        if (x.substring(0, 3) === 'EUR' && x.length < 5) {
          month_columns.push(key.substring(0, key.length - 2));
        }
      }
    });
    //  month_columns = [ 'M', 'O', 'Q', 'S', 'U', 'W' ]   columns of months in "Detailed view"
    const sumRow = 286; // sum of all measures spent budget per month in this row in "Detailed view"
    const monthlySpendings = month_columns.map((month, index) => {
      return Math.round(detailes_object[month + '' + sumRow]['v'] * 100) / 100;
    });
    const approvedBudgetPerMonth =
      Math.round((totalApprovedBudget / month_columns.length) * 100) / 100;
    const year = new Date().getFullYear(); // TO DO: year is currently assumed to be this year (Datetime). Improve by parsing from file
    const newBudget = new this.budgetModel({
      monthlySpendings,
      approvedBudgetPerMonth,
      year,
    });
    newBudget.save().then((result) => {
      console.log('budget saved');
    });
    const excelSheet = await this.sheetModel.findOne({
      name: fileNames.main_file,
    });
    excelSheet
      .update({ totalBudget: totalApprovedBudget })
      .then((result) => {
        console.log(result);
      })
      .catch((e) => {
        console.log(e);
      });
    return 'budget parsed';
  }
}
