using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace ErrorManager {
    class ErrorListBook {

        public ErrorListBook(){

        }

        public string Name { get; set; }
        public List<string> ErrorList { get; set; }
        public string SekkeishoPath { get; set; }

        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        Excel.Range range;

        /// <summary>
        /// Read Excel and get data at column 8
        /// If kind of Error not contian already in ErrorList, add to ErrorList
        /// </summary>
        /// <param name="Files">List of Error Files</param>
        /// <param name="path">Error Folder path to use as Export Path</param>
        /// <param name="background">BackgroundWorker</param>
        public Boolean GetErrorListFormExcel(List<string> Files, string path, BackgroundWorker background) {
            List<ErrorListBook> list = new List<ErrorListBook>();
            List<string> colList = new List<string>();

            Excel.Application ExcelApp = new Excel.Application();
            int progress = 1;
            
            foreach (string file in Files){
                
                if(!background.CancellationPending){
                    workbook = ExcelApp.Workbooks.Open(file);
                    worksheet = workbook.Sheets[1];
                    range = worksheet.UsedRange;
                    int rowCount = range.Rows.Count;
                    int dataColumn = 8; //Column No. to Read

                    ErrorListBook book = new ErrorListBook();
                    book.ErrorList = new List<string>();
                    book.Name = range.Cells[2, 1].Value2;
                    book.SekkeishoPath = range.Cells[2, 4].Value2 + "\\"+ range.Cells[2, 3].Value2;

                    for (int i = 2; i <= rowCount; i++) {
                        string cellValue = range.Cells[i, dataColumn].Value2;
                        if (range.Cells[i, dataColumn] != null && cellValue != null) {
                            // If kind of error not contian already in ErrorList, add to ErrorList
                            if(cellValue.Equals("表種")){
                                cellValue = cellValue + "," + range.Cells[i, dataColumn + 1].Value2 + " << " + range.Cells[i, dataColumn + 2].Value2;
                            }
                            if (!book.ErrorList.Contains(cellValue)) {
                                book.ErrorList.Add(cellValue);
                            }

                            //Add data for column name
                            if (cellValue.Contains("表種")) {
                                if (!colList.Contains(cellValue.Split(',')[1])) {
                                    colList.Add(cellValue.Split(',')[1]);
                                }
                            } else if (!colList.Contains(cellValue)) {
                                colList.Add(cellValue);
                            }
                        }
                    }
                    list.Add(book);
                    background.ReportProgress(progress++);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //release com objects to fully kill excel process from running in the background
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(worksheet);

                    //close and release
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }else{
                    //quit and release
                    ExcelApp.Quit();
                    Marshal.ReleaseComObject(ExcelApp);
                    return true;
                }
            }
            colList.Sort();
            // Create Excel file with data in list
            ExportDataSetToExcel(list, ExcelApp, path, colList);
            return false;
        }

        Excel.Range nrange;

        /// <summary>
        /// Create Excel file with data in list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="ExcelApp"></param>
        /// <param name="path">Export Paht</param>
        public void ExportDataSetToExcel(List<ErrorListBook> list, Excel.Application ExcelApp, string path, List<string> colList) {
            try {
                ExcelApp.Visible = false;
                ExcelApp.DisplayAlerts = false;
                workbook = ExcelApp.Workbooks.Add();
                
                //Create an Excel workbook instance and open it from the predefined location
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "ErrorManger";

                //List<string> colList = new List<string>();
                string ExportPath = path + @"/ErrorManger";

                for (int j = 0; j < list.Count; j++) {
                    worksheet.Cells[j + 2, 1] = list[j].Name;
                    worksheet.Cells[j + 2, 2] = list[j].SekkeishoPath;
                    foreach(string item in list[j].ErrorList){
                        //Add data for Row data
                        if(item.Contains("表種")){
                            worksheet.Cells[j + 2, colList.IndexOf(item.Split(',')[1]) + 3] = item.Split(',')[1];
                        } else{
                            worksheet.Cells[j + 2, colList.IndexOf(item) + 3] = "〇";
                        }
                        
                    }   
                }
                
                worksheet.Cells[1, 1] = "テスト名";
                worksheet.Cells[1, 2] = "取込ファイル";
                for (int i=0; i<colList.Count; i++){
                    if(colList[i].Contains("<<")){
                        worksheet.Cells[1, i + 3] = colList[i];
                    } else{
                        worksheet.Cells[1, i + 3] = colList[i];
                    }
                }

                //Setting Output File Design Format
                nrange = worksheet.UsedRange;
                nrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                nrange = worksheet.Columns["A"];
                nrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                nrange.Columns.AutoFit();
                nrange = worksheet.Columns["B"];
                nrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignFill;
                nrange = worksheet.Rows[1];
                nrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                nrange.Columns.AutoFit();
                
                //nrange.AutoFilter(1,"<>",XlAutoFilterOperator.xlOr, true);
                ExcelApp.ActiveWindow.SplitRow = 1;
                ExcelApp.ActiveWindow.SplitColumn = 2;
                ExcelApp.ActiveWindow.FreezePanes = true;


                if(!Directory.Exists(ExportPath)){
                    Directory.CreateDirectory(ExportPath);
                }
                workbook.SaveAs(ExportPath + @"/ErrorManager_"+ DateTime.Now.ToString("yyyy-dd-M_HH-mm-ss") +".xlsx");

            } catch (Exception e) {
                MessageBox.Show(e.Message, e.GetType().ToString());
            }finally{
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(worksheet);

                //close and release
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
                //quit and release
                ExcelApp.Quit();
                Marshal.ReleaseComObject(ExcelApp);
            }
        }
    }
}
