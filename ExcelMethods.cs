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
                    _workbook.ActiveSheet.Clear();
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

        internal void ToExcel(int count, Dictionary<string, List<string>>.KeyCollection keys, Dictionary<string, List<string>>.ValueCollection values)
        {

        }
        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
