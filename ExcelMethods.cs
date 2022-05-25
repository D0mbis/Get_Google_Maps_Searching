using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Selenium
{
    internal class ExcelMethods : IDisposable
    {
        private Excel.Application _excel;
        private string path = @"D:\\Programming_study\\Selenium\\Result\\List.xls";
        private Excel.Workbook _workbook;
        public ExcelMethods()
        {
            _excel = new Excel.Application();
        }

        internal void Open()
        {
            try
            {
                if (File.Exists(path))
                {
                    _workbook = _excel.Workbooks.Open(path);
                    _workbook.Worksheets[1].Cells.ClearContents();
                    _workbook.ActiveSheet.Name = "Список результатів";
                }
                else
                {
                    using (FileStream fs = new FileStream(path, FileMode.Create)) { }
                    _workbook = _excel.Workbooks.Open(path);
                    _workbook.ActiveSheet.Name = "Список результатів";
                }
            }
            catch
            {
                Dispose();
                MessageBox.Show("Не вдалось запустити Excel, можливо версія несумісна");
            }
        }

        internal void ToExcel(int row, int column, object data)
        {
            _workbook.ActiveSheet.Cells[row, column] = data;
        }
        public void Dispose()
        {
            try
            {
                //_workbook.SaveAs(FileFormat: XlFileFormat.xlAddIn8);
                _workbook.Save();
                _workbook.Close();
                _excel.Quit();
                MessageBox.Show("Excel завершила свою роботу!");
            }
            catch
            {
                MessageBox.Show("Не вдається завершити роботу Excel!");
            }
        }
    }
}
