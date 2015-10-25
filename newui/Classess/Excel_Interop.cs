using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using VNS.Libs;
using DataTable = System.Data.DataTable;

namespace VNS.HIS.UI.Classess
{
    public class  Excel_Interop
    {
        public static DataTable LoadDataFromFileExcelToDataTable(string Path)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            var dtaTable = new DataTable();
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(Path, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value);

                var ws =(Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets.Item["Data"];

                Range excelRange = ws.UsedRange; //gives the used cells in sheet

                ws = null; // now No need of this so should expire.

                //Reading Excel file.               
                var valueArray = (object[,])excelRange.Value[Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault];

                excelRange = null; // you don't need to do any more Interop. Now No need of this so should expire.
                dtaTable = ProcessObjects(valueArray);
                return dtaTable;
            }
            catch (Exception ex)
            {
                Utility.ShowMsg(ex.Message);
                return null;
            }
            finally
            {
                #region Clean Up

                if (workbook != null)
                {
                    #region Clean Up Close the workbook and release all the memory.

                    workbook.Close(false, Path, Missing.Value);
                    Marshal.ReleaseComObject(workbook);

                    #endregion
                }
                workbook = null;

                if (excelApp != null)
                {
                    excelApp.Quit();
                }
                excelApp = null;

                #endregion
            }

        }
        public static DataTable ProcessObjects(object[,] valueArray)
        {
            var dt = new System.Data.DataTable();

            #region Get the COLUMN names

            for (int k = 1; k <= valueArray.GetLength(1); k++)
            {
                dt.Columns.Add((string)valueArray[1, k]);  //add columns to the data table.
            }
            #endregion

            #region Load Excel SHEET DATA into data table

            object[] singleDValue = new object[valueArray.GetLength(1)];
            //value array first row contains column names. so loop starts from 2 instead of 1
            for (int i = 2; i <= valueArray.GetLength(0); i++)
            {
                for (int j = 0; j < valueArray.GetLength(1); j++)
                {
                    if (valueArray[i, j + 1] != null)
                    {
                        singleDValue[j] = valueArray[i, j + 1].ToString();
                    }
                    else
                    {
                        singleDValue[j] = valueArray[i, j + 1];
                    }
                }
                dt.LoadDataRow(singleDValue, System.Data.LoadOption.PreserveChanges);
            }
            #endregion
            return (dt);
        }

    }
}
