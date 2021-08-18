import { Injectable } from '@nestjs/common';
import * as fs from "fs";
import { InjectModel } from '@nestjs/mongoose';
import { Sheet, SheetSchema } from '../schemas/sheet.schema';
import { Measure, MeasureSchema } from '../schemas/measure.schema';
import { Artefact, ArtefactSchema } from '../schemas/artefact.schema';
import { Budget, BudgetSchema } from '../schemas/budget.schema';
import { Model } from 'mongoose';
import { resolve } from "path";
import * as XLSX from "xlsx";


@Injectable()
export class XlsxParserService {
  fileName: string = 'test_data.xlsx'
  kpi_file_1: string = 'KPI-report_1.xlsx'
  kpi_file_2: string = 'KPI-report_2.xlsx'
  budget_file: string = 'budget_report.xlsx'

  constructor(
    @InjectModel('Artefact') private artefactModel: Model<Artefact>,
    @InjectModel('Measure') private measureModel: Model<Measure>,
    @InjectModel('Sheet') private sheetModel: Model<Sheet>,
    @InjectModel('Budget') private budgetModel: Model<Budget>,
  ) { }


  /*
    async getAllArtefacts(): Promise<string> {
      const result = await this.artefactModel.find()
      console.log(result)
      return JSON.stringify(result)
    }
  */

  async getArtefactsOfMeasure(measureID: string): Promise<string> {
    const measure = await this.measureModel.findById(measureID)
    if (measure) {
      const populatedMeasure = await measure.populate("artefacts").execPopulate()
      return JSON.stringify(populatedMeasure.artefacts)
    } else {
      return JSON.stringify("Error")
    }
  }


  async getAllMeasures(): Promise<string> {
    const result = await this.measureModel.find()
    console.log(result)
    return JSON.stringify(result)
  }


  async getOverview(): Promise<string> {
    const excelSheet = await this.sheetModel.findOne({ name: this.fileName })   // TO DO: safeguard for duplicates
    console.log(excelSheet)
    return JSON.stringify(excelSheet)
  }


  async getBudget(): Promise<string> {
    const excelSheet = await this.budgetModel.findOne()   // TO DO: safeguard for duplicates
    console.log(excelSheet)
    return JSON.stringify(excelSheet)
  }


  // ---------  parsing functions  -----------------------------

  async createOverview(): Promise<string> {
    const excelSheet = await this.sheetModel.findOne({ name: this.fileName })   // TO DO: safeguard for duplicates
    if (excelSheet) {
      const numberOfMeasures = excelSheet.measures.length
      console.log(numberOfMeasures)

      const workbook = XLSX.readFile(resolve(__dirname, this.fileName))
      const overview_object = workbook.Sheets["Status Overview"]
      //    console.log[overview_object]

      // ------total budget
      let totalBudget = 0;
      let allBudgetsOfMeasures = []                     // <-----------------  [{measureName, budget}, ...]
      Object.keys(overview_object).filter(key => {
        if (key.includes("I")) {
          const row = parseInt(key.substring(1))
          if (row > 4) {
            const measureName = overview_object["D" + row]["v"]
            //         console.log(overview_object[key])
            const budgetAsString = overview_object[key]["v"]
            const budget = parseInt(budgetAsString.substring(0, budgetAsString.indexOf("k"))) * 1000
            totalBudget += budget
            const currentBudget = {
              [measureName]: budget
            }
            allBudgetsOfMeasures.push(currentBudget)
          }
        }
      })
      // ------overall status
      const all_risksBudgetsAndArtefacts = []
      Object.keys(overview_object).filter(key => {
        if (key.includes("P")) {
          const row = parseInt(key.substring(1))
          if (row > 4) {
            all_risksBudgetsAndArtefacts.push(overview_object[key]["v"])
          }
        }
      })
      Object.keys(overview_object).filter(key => {
        if (key.includes("Q")) {
          const row = parseInt(key.substring(1))
          if (row > 4) {
            all_risksBudgetsAndArtefacts.push(overview_object[key]["v"])
          }
        }
      })
      Object.keys(overview_object).filter(key => {
        if (key.includes("R")) {
          const row = parseInt(key.substring(1))
          if (row > 4) {
            all_risksBudgetsAndArtefacts.push(overview_object[key]["v"])
          }
        }
      })
      //    console.log(all_risksBudgetsAndArtefacts)
      let overallStatus = 0
      all_risksBudgetsAndArtefacts.map(a => {
        if (a > overallStatus) { overallStatus = a }
      })

      // ------Progress Overview:  sum over measures(avgProgress * measureBudget) / totalBudget
      // get all measures, get artefacts of each measure    // why only measure_id notfull object ????  => compare schemas artefacts and measures
      // console.log(allBudgetsOfMeasures)
      let sum_avgProgressTimesBudgetOfMEasures = 0

      for (let m = 0; m < excelSheet.measures.length; m++) {
        const measure = await (await this.measureModel.findById(excelSheet.measures[m])).populate("artefacts").execPopulate()
        let avgProgressOfArtefacts = 0
        measure.artefacts.map(art => {
          avgProgressOfArtefacts += art.progress
        })
        avgProgressOfArtefacts = avgProgressOfArtefacts / measure.artefacts.length
        let temp = 0
        let budgetOFThisMEasure
        allBudgetsOfMeasures.map(item => {
          if (item[measure.title]) {
            //    console.log(measure.title + "    "  +  item[measure.title])
            temp = item[measure.title] * avgProgressOfArtefacts
            budgetOFThisMEasure = item[measure.title]
          }
        })
        console.log(measure.title + "  " + avgProgressOfArtefacts + " * " + budgetOFThisMEasure + ":  " + temp)
        sum_avgProgressTimesBudgetOfMEasures += temp
      }
      const progressOverviewBar_result = sum_avgProgressTimesBudgetOfMEasures / totalBudget

      // ------KPI Progress 
      const KPIprogressOfAllMeasures = this.getKPIProgressData('kpi_progress.xlsx')
      let sum = 0
      KPIprogressOfAllMeasures.map(item => {    //  [{measurename: ..., kpiProgress: ...}, ...]
        let thisBudget = 0
        allBudgetsOfMeasures.map(budgetOfMeasure => {
          if (budgetOfMeasure[item.measureName]) {
            const temp = item.progress * budgetOfMeasure[item.measureName]
            sum += temp
            console.log(item.measureName + "   " + item.progress + " *  " + budgetOfMeasure[item.measureName] + ":  " + temp)
          }
        })
      })
      const KPIProgress_result = sum / totalBudget
      excelSheet.update({
        totalBudget: totalBudget,
        overallStatus: overallStatus,
        progress: Math.round(progressOverviewBar_result * 100) / 100,
        kpiProgress: Math.round(KPIProgress_result * 100) / 100,
      }).then(result => {
        console.log("updated")
      })
      console.log("numberOfMeasures   " + numberOfMeasures)
      console.log("totalBudget   " + totalBudget)
      console.log("overallStatus   " + overallStatus)
      console.log("progressOverviewBar_result   " + progressOverviewBar_result)
      console.log("KPIProgress_result   " + KPIProgress_result)
      //######################
    }
    const result = "moin"
    //console.log(result)
    return JSON.stringify(result)
  }




  // aux function for createOverview()
  getKPIProgressData(kpiFile: string): any {
    const workbook = XLSX.readFile(resolve(__dirname, kpiFile))
    const overview_object = workbook.Sheets["Plan view"]
    //  console.log(overview_object)
    // D:measure name, G current progress, H target progress
    // TO DO: get number of rows programmatically    ---   22
    const numberOfRows = 22
    const result = []
    for (let i = 1; i <= numberOfRows; i++) {
      const key_measureName = "D" + (4 + i)       // first enry at row 4
      const key_actualProgress = "G" + (4 + i)
      const key_targetProgress = "H" + (4 + i)
      result.push({
        measureName: overview_object[key_measureName]["v"],
        progress: Math.round(((overview_object[key_actualProgress]["v"] / overview_object[key_targetProgress]["v"])) * 100) / 100
      })
    }
    //  console.log(result)
    return result
  }
  //------------------------------------------------






  // parse and save measures and corresponding artefacts
  parse(): string {
    // create Sheet Table
    const newSheet = {
      name: this.fileName,
    }
    const excelFile = new this.sheetModel(newSheet)
    excelFile.save()
      .then(newlySavedExcelSheet => {
        console.log(newlySavedExcelSheet)

        const workbook = XLSX.readFile(resolve(__dirname, this.fileName))
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
                const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName])
                const artefacts = this.getArtefacts_fromLinesArray(data)

                const savedArtefact_IDs = []
                artefacts.map(art => {
                  const toSave = {
                    id: art["__EMPTY_1"],
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
  getArtefacts_fromLinesArray(sheet: Object[]): Object[] {
    return sheet.filter(line => {
      const firstKey = Object.keys(line)[0]
      if (firstKey === "__EMPTY_1") {
        const firstItem = line[firstKey]
        try {
          if (parseInt(firstItem) < 10 && Object.keys(line).length > 2) {
            return line
          }
        } catch (e) {
          console.log("-")
        }
      }
    })
  }






  // adds status info to measures 
  parse_overview(): string {
    const workbook = XLSX.readFile(resolve(__dirname, this.fileName))
    // parse overview
    const overview_object = workbook.Sheets["Status Overview"]
    const risks = []                                                    // RESTRUCTURE!!!!!!!
    Object.keys(overview_object).filter(key => {
      if (key.includes("P")) {
        const row = parseInt(key.substring(1))
        if (row > 4) {
          risks.push({
            row: row,
            risk: overview_object[key]["v"],
          })
        }
      }
    })
    const budgets = []
    Object.keys(overview_object).filter(key => {
      if (key.includes("Q")) {
        const row = parseInt(key.substring(1))
        if (row > 4) {
          budgets.push({
            row: row,
            budget: overview_object[key]["v"],
          })
        }
      }
    })
    const artefacts = []
    Object.keys(overview_object).filter(key => {
      if (key.includes("R")) {
        const row = parseInt(key.substring(1))
        if (row > 4) {
          artefacts.push({
            row: row,
            artefact: overview_object[key]["v"],
          })
        }
      }
    })
    const result = []
    Object.keys(overview_object)
      .filter(async key => {
        if (key.includes("D")) {
          const row = parseInt(key.substring(1))
          if (row > 4) {
            const addToResult = {
              row: row,
              name: overview_object[key]["h"],
              risk: risks[row - 5]["risk"],
              budget: budgets[row - 5]["budget"],
              artefact: artefacts[row - 5]["artefact"],
            }
            await this.measureModel.updateOne({ title: addToResult.name }, {
              risk: addToResult.risk,
              budget: addToResult.budget,
              artefact: addToResult.artefact

            });
            result.push(addToResult)
          }
        }
      })
    console.log(result)
    return JSON.stringify(result)
  }








  async parseKPI(): Promise<string> {
    const workbook = XLSX.readFile(resolve(__dirname, this.kpi_file_2))
    const overview_object = workbook.Sheets["Plan view"]
    // get row number of measures. measure names in column "D"
    const rowsOfMeasures = []
    Object.keys(overview_object).filter(key => {
      if (key.includes("D")) {
        const row = parseInt(key.substring(1))
        if (row > 4) {
          rowsOfMeasures.push({
            measureName: overview_object[key]["v"],
            row: row,
          })
        }
      }
    })
    //     console.log(rowsOfMeasures)
    // get measures from DB
    const measures = await this.measureModel.find()
    measures.map(measure => {
      let rowOfThisMeasure;
      for (let i = 0; i < rowsOfMeasures.length; i++) {
        if (rowsOfMeasures[i].measureName === measure.title) {
          //   console.log(measure.title + "    " + rowsOfMeasures[i].measureName)
          rowOfThisMeasure = rowsOfMeasures[i].row
        }
      }
      console.log(measure.title + "    " + rowOfThisMeasure)
      // from row get actual, target, plan of last month
      // G, H and J ???
      let actuals;
      let target;
      let lastPlan;
      Object.keys(overview_object).filter(key => {
        if (key.includes("G")) {
          const row = parseInt(key.substring(1))
          if (row == rowOfThisMeasure) {
            actuals = overview_object[key]["v"]
          }
        }
        if (key.includes("H")) {
          const row = parseInt(key.substring(1))
          if (row == rowOfThisMeasure) {
            target = overview_object[key]["v"]
          }
        }
        if (key.includes("L")) {                        // TO DO: clarify which last month
          const row = parseInt(key.substring(1))
          if (row == rowOfThisMeasure) {
            lastPlan = overview_object[key]["v"]
          }
        }
      })
      console.log(measure.title + "  " + actuals + "  " + target + "  " + lastPlan)
      let kpiProgressOfThisMeasure;
      if (actuals < lastPlan) {
        kpiProgressOfThisMeasure = 0                      // behind schedule
      } else if (lastPlan <= actuals && actuals < target) {
        kpiProgressOfThisMeasure = 1                      // on schedule
      } else if (actuals >= target) {
        kpiProgressOfThisMeasure = 2                      // finished
      }
      measure.update({
        kpiProgress: kpiProgressOfThisMeasure
      }).then(result => {
        console.log("updated")
      })
    })
    return "ok"
  }







  async parseBudgetMonths(): Promise<string> {
    const workbook = XLSX.readFile(resolve(__dirname, this.budget_file))
    const overview_object = workbook.Sheets["1. Overview"]
    const detailes_object = workbook.Sheets["2. Detailed view"]
    // M > 5,  D measure names
    // M 29 grand total approved budget
    const totalApprovedBudget = overview_object["M29"]["v"]

    // get month columns EUR1 EUR2 .... row 12
    const month_columns = []
    Object.keys(detailes_object).map(key => {
      //       const test = key.substring(key.length - 2)
      //       console.log(test)
      const tmp = key.replace(/^[A-Z]/, "_");
      const split = tmp.split("_")
      const target = split[split.length - 1]
      if (parseInt(target) == 12) {
        const x = detailes_object[key]["v"]
        if (x.substring(0, 3) === "EUR" && x.length < 5) {
          //   console.log(x)
          month_columns.push(key.substring(0, key.length - 2))
        }
      }
    })
    //  month_columns = [ 'M', 'O', 'Q', 'S', 'U', 'W' ]   columns of months in "Detailed view"
    const sumRow = 286    // sum of all measures spent budget per month in this row in "Detailed view"
    // get sum of spent berson / month for all measures combined
    const monthNames = ["jan", "feb", "mar", "apr", "mai", "jun", "jul", "auf", "sep", "oct", "nov", "dec"]
    const monthlySpendings = month_columns.map((month, index) => {
      return Math.round(detailes_object[month + "" + sumRow]["v"] * 100) / 100
      /*      return {
              [monthNames[index]]: Math.round(detailes_object[month + "" + sumRow]["v"] * 100) / 100
            }  */
    });
    console.log(monthlySpendings)


    const approvedBudgetPerMonth = Math.round((totalApprovedBudget / month_columns.length) * 100) / 100
    const year = new Date().getFullYear()  // TO DO: year is currently assumed to be this year (Datetime). Improve by parsing from file
    const newBudget = new this.budgetModel({
      monthlySpendings,
      approvedBudgetPerMonth,
      year
    })
    newBudget.save()
      .then(result => {
        console.log("budget saved")
      }
      )
    // add totalapprovedbudget to Sheet Table
    const excelSheet = await this.sheetModel.findOne({ name: this.fileName })   // TO DO: safeguard for duplicates
    excelSheet.update({ totalBudget: totalApprovedBudget })
      .then(result => {
        console.log(result)
      })
      .catch(e => {
        console.log(e)
      })

    return JSON.stringify("budget parsed")
  }
























}
