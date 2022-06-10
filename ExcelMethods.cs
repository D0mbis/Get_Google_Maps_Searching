using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Selenium
{
    internal class ExcelMethods : IDisposable
    {
        private Excel.Application _excel = null;
        private Excel.Workbook _workbook = null;
        private Excel.Worksheet _sheet = null;
        private string path = null;
        public ExcelMethods()
        {
            _excel = new Excel.Application();
        }
        public void SaveAs()
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Книга Excel 97-2003 (*.xls) | *.xls";
            if (saveDialog.ShowDialog() == true)
            {
                path = saveDialog.FileName;
                //_workbook.SaveAs(path);
            }
        }

        internal void Open()
        {
            if (path != null)
            {
                _workbook = _excel.Workbooks.Open(path);
                _workbook.Worksheets[1].Cells.ClearContents();
            }
            else
            {
                try
                {
                    _workbook = _excel.Workbooks.Add();
                    _sheet = _workbook.Sheets.Add();
                    _sheet.Name = "Results List";
                }
                catch
                {
                    Dispose();
                    MessageBox.Show("Excel does not start, the version may be incompatible");
                }
            }
        }

        internal void ToExcel(int row, int column, object data)
        {
            _workbook.ActiveSheet.Cells[row, column] = data;
        }

        internal void Save()
        {
            try
            {
                _workbook.SaveAs(path);
                _workbook.Close();
                _excel.Quit();
                MessageBox.Show("Results was saved!");
            }
            catch
            {
                MessageBox.Show("You need to set the path to the file with the results!");
                SaveAs();
                if (path != null)
                {
                    _workbook.SaveAs(path);
                    _workbook.Close();
                    _excel.Quit();
                    MessageBox.Show("Results was saved!");
                }
                else
                    MessageBox.Show("Excel don`t finished work!");
            }
        }
        public void Dispose()
        {

        }
    }
}
