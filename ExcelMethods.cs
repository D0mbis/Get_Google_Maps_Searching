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
        private string path = @"D:\\Programming_study\\Selenium\\Result\\list.xls";
        private Excel.Workbook _workbook;
        private string needSave;

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
                    //_workbook.ActiveSheet.Cells.ClearContents(path);
                    //_workbook.Sheets.Delete(); ??
                    //File.Delete(path);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add(path);
                    needSave = path;
                }
            }
            catch
            {
                MessageBox.Show("Не вдалось запустити Excel, можливо версія несумісна");
            }
        }

        internal void ToExcel(int row, int column, object data)
        {
            _workbook.ActiveSheet.Cells[row, column] = data;
        }

        internal void Save()
        {
            _workbook.Save();
        }
        public void Dispose()
        {
            try
            {
                _workbook.Close();
            }
            catch
            {
                MessageBox.Show("Не вдається завершити роботу Excel!");
            }
        }
    }
}
