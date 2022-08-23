using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace CNO.BPA.GenerateXLS
{
    class XLSBuilder
    {
        public void createXLS(DataSet dsBatchDetails)
        {
            System.Data.DataTable dtBatchItemDetails = new System.Data.DataTable();
            dtBatchItemDetails = dsBatchDetails.Tables[0];

            object misValue = System.Reflection.Missing.Value;

            string strExcelFileDir = string.Empty;
            string strExcelFilePath = string.Empty;
                        
            //Diretory Path
            strExcelFileDir = BatchDetail.DefaultXLSPath + "\\CASH_NON_FORM\\" + BatchDetail.Department + "\\" + BatchDetail.BatchNo;

            //Check if the directory exists and create directory if it doesnot exist                
            if (!Directory.Exists(strExcelFileDir))
            {
                Directory.CreateDirectory(strExcelFileDir);
            }

            //Excel Filepath
            strExcelFilePath = strExcelFileDir + "\\" + BatchDetail.BatchNo;

            try
            {
                if (dtBatchItemDetails == null || dtBatchItemDetails.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                Excel.Application excelApp = new Excel.ApplicationClass();
                Excel.Workbook workbook;
                workbook = excelApp.Workbooks.Add(misValue);

                // single worksheet
                Excel._Worksheet workSheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
                workSheet.Name = BatchDetail.BatchNo;

                // column headings
                for (int i = 0; i < dtBatchItemDetails.Columns.Count; i++)
                {
                    workSheet.Cells[1, (i + 1)] = dtBatchItemDetails.Columns[i].ColumnName;
                }

                // rows
                for (int i = 0; i < dtBatchItemDetails.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (int j = 0; j < dtBatchItemDetails.Columns.Count; j++)
                    {
                        workSheet.Cells[(i + 2), (j + 1)] = "\'" + dtBatchItemDetails.Rows[i][j];
                    }
                }

                //check filepath
                if (strExcelFilePath != null && strExcelFilePath != "")
                {
                    try
                    {
                        workbook.SaveAs(strExcelFilePath + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        workbook.Close(true, misValue, misValue);
                        excelApp.Quit();

                        //Release objects
                        releaseObject(workSheet);
                        releaseObject(workbook);
                        releaseObject(excelApp);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else
                {
                    workbook.Close(true, misValue, misValue);
                    excelApp.Quit();                      

                    //excelApp.Visible = true;
                    //log.Error("Verify the excel path");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }          
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;                
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
