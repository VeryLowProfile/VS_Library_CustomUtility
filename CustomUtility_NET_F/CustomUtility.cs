using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CustomUtility_NET_F
{
    #region ChildForm Calss
    //Class to manage a child form on a panel
    public static class ChildForm
    {
        public static void OpenChildFormOnPanel(Form newForm, Panel panel)
        {
            try
            {
                panel.Controls.Clear();
                newForm.TopLevel = false;
                newForm.FormBorderStyle = FormBorderStyle.None;
                newForm.Dock = DockStyle.Fill;
                panel.Controls.Add(newForm);
                panel.Tag = newForm;
                newForm.BringToFront();
                newForm.Show();
            }
            catch (Exception ex)
            {
                //display error message
                MessageBox.Show("Exception: Class ChildForm - Metod OpenChildFormOnPanel" + ex.Message);
            }
        }
    }
    #endregion

    #region ExcelDataTable
    //Calss to import/export from/to DataTable/Excel Sheet
    public static class ExcelDataTable
    {
        public static DataTable ImportExcelToDataTable(string filePath, int sheetIndex)
        {
            //New Data Table
            DataTable dataTable = new DataTable();

            //New Data Row
            DataRow row;

            try
            {
                //Open New Xcel application
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(filePath);
                Excel.Worksheet worksheet = workbook.Sheets[sheetIndex + 1];
                Excel.Range range = worksheet.UsedRange;
                excel.DisplayAlerts = false;

                //Data Table create new coloumns
                for (int i = 1; i <= range.Columns.Count; i++)
                {
                    dataTable.Columns.Add(range.Cells[1, i].Value2);
                }

                //Fill Data Table With New Rows From Excel File
                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    row = dataTable.NewRow();
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        {
                            row[j - 1] = range.Cells[i, j].Value2;

                        }
                        else
                        {
                            row[j - 1] = " ";
                        }
                    }
                    dataTable.Rows.Add(row);
                }

                workbook.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                //display error message
                MessageBox.Show("Exception: Class ExcelDataTable - Metod ImportExcelFromDataTable " + ex.Message);
            }

            return dataTable;
        }

        public static void ExportDataTableToExcel(DataTable dataTable, string filePath, string sheetName)
        {
            try
            {
                //Open Excel and create new sheet
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = sheetName;
                excel.DisplayAlerts = false;

                //Set Coloumn Header
                for (int i = 0; i <= dataTable.Columns.Count - 1; i++)
                {
                    worksheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName.ToString();
                }

                //Export Table
                for (int i = 0; i <= dataTable.Rows.Count - 1; i++)
                {
                    for (int j = 0; j <= dataTable.Columns.Count - 1; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j];
                    }
                }

                workbook.SaveAs(filePath);

                workbook.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                //display error message
                MessageBox.Show("Exception: Class ExcelDataTable - Metod ExportDataTableToExcel " + ex.Message);
            }
        }

        public static List<string> GetSheetsCollection(string filePath)
        {
            //New string List
            List<string> list = new List<string>();

            try
            {

                //Open New Xcel application
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(filePath);

                //Fill Combobox
                foreach (Excel.Worksheet worksheet in excel.Worksheets)
                {
                    list.Add(worksheet.Name);
                }

                workbook.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                //display error message
                MessageBox.Show("Exception: " + ex.Message);
            }

            return list;

        }

    }

    #endregion
}
