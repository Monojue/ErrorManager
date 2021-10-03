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
        public void GetErrorListFormExcel(List<string> Files, string path, BackgroundWorker background) {
            List<ErrorListBook> list = new List<ErrorListBook>();

            Excel.Application ExcelApp = new Excel.Application();
            int progress = 1;
            
            foreach (string file in Files){
                
                if(!background.CancellationPending){
                    workbook = ExcelApp.Workbooks.Open(file);
                    worksheet = workbook.Sheets[1];
                    range = worksheet.UsedRange;
                    int rowCount = range.Rows.Count;
                    int dataColumn = 8; //Column No to Read

                    ErrorListBook book = new ErrorListBook();
                    book.ErrorList = new List<string>();
                    book.Name = range.Cells[2, 1].Value2;
                    for (int i = 2; i <= rowCount; i++) {
                        string cellValue = range.Cells[i, dataColumn].Value2;
                        if (range.Cells[i, dataColumn] != null && cellValue != null) {
                            // If kind of error not contian already in ErrorList, add to ErrorList
                            if (!book.ErrorList.Contains(cellValue)) {
                                if(cellValue.Equals("表種")){
                                    book.ErrorList.Add(cellValue +","+ range.Cells[i, dataColumn+ 1].Value2 + " << "+ range.Cells[i, dataColumn + 2].Value2);
                                } else{
                                    book.ErrorList.Add(cellValue);
                                }
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
                }
            }
            if (background.CancellationPending){
                //quit and release
                ExcelApp.Quit();
                Marshal.ReleaseComObject(ExcelApp);
            } else{
                // Create Excel file with data in list
                ExportDataSetToExcel(list, ExcelApp, path);
            }
        }

        Excel.Range nrange;

        /// <summary>
        /// Create Excel file with data in list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="ExcelApp"></param>
        /// <param name="path">Export Paht</param>
        public void ExportDataSetToExcel(List<ErrorListBook> list, Excel.Application ExcelApp, string path) {
            try {
                ExcelApp.Visible = false;
                ExcelApp.DisplayAlerts = false;
                workbook = ExcelApp.Workbooks.Add();
                
                //Create an Excel workbook instance and open it from the predefined location
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "ErrorManger";

                List<string> colList = new List<string>();
                string ExportPath = path + @"/ErrorManger";

                for (int j = 0; j < list.Count; j++) {
                    worksheet.Cells[j + 2, 1] = list[j].Name;
                    foreach(string item in list[j].ErrorList){
                        //Add data for column name
                        if(item.Contains("表種") && !colList.Contains("表種")){
                            colList.Add("表種");
                        }
                        else if(!colList.Contains(item) && !item.Contains("表種")) {
                            colList.Add(item);
                        }

                        //Add data for Row data
                        if(item.Contains("表種")){
                            worksheet.Cells[j + 2, colList.IndexOf("表種") + 2] = item.Split(',')[1];
                        } else{
                            worksheet.Cells[j + 2, colList.IndexOf(item) + 2] = "〇";
                        }
                        
                    }   
                }
                
                
                worksheet.Cells[1, 1] = "テスト名";
                for(int i=0; i<colList.Count; i++){
                    worksheet.Cells[1, i+2] = colList[i];
                }
                nrange = worksheet.UsedRange;
                nrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                nrange.Columns.AutoFit();
                nrange.AutoFilter(1,"<>",XlAutoFilterOperator.xlOr, true);
                ExcelApp.ActiveWindow.SplitRow = 1;
                ExcelApp.ActiveWindow.SplitColumn = 1;
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
