const distPath = __dirname; // ...pmo/pmo-backend/dist
const rootPath_array = distPath.split('/');
rootPath_array.pop();
const rootPath = rootPath_array.join('/') + '/src/';
const xlsx_file_dir = rootPath + 'xlsx_files/';

export const fileNames = {
  xlsx_file_dir,
  main_file: 'test_data.xlsx',
  kpi_file_1: 'KPI-report_1.xlsx',
  kpi_file_2: 'KPI-report_2.xlsx',
  budget_file: 'budget_report.xlsx',
  status_report: 'status_report.xlsx'
};
