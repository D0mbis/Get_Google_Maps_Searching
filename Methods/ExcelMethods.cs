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
        private string path;

        public ExcelMethods()
        {
            _excel = new Excel.Application();
            using (StreamReader stream = new StreamReader("path.txt")) { path = stream.ReadToEnd(); }
        }

        public string SaveAs()
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Книга Excel 97-2003 (*.xlsx) | *.xlsx";
                if (saveDialog.ShowDialog() == true)
                {
                    path = saveDialog.FileName.ToString();
                }
                return path;
            }
            catch
            {
                return null;
            }
        }

        internal void Open()
        {
            if (File.Exists(path))
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
                    MessageBox.Show("Error! \n Excel does not start, the version may be incompatible");
                }
            }
        }

        internal void ToExcel(int row, int column, object data)
        {
            _workbook.ActiveSheet.Cells[row, column] = data;
        }

        public string Save()
        {
            try
            {
                if (File.Exists(path))
                {
                    _workbook.SaveAs(path);
                    using (StreamWriter stream = new StreamWriter("path.txt")) { stream.Write(path); }
                    MessageBox.Show("Results was saved successfully!");
                }
                else
                {
                    MessageBox.Show("Save error! \n You need to correctly set the path to save the file. Try again");
                    SaveAs();
                    _workbook.SaveAs(path);
                    using (StreamWriter stream = new StreamWriter("path.txt")) { stream.Write(path); }
                    MessageBox.Show("Results was saved successfully!");

                }
            }
            catch
            {
                //MessageBox.Show("You need to set the path to the file with the results!");
                do
                {
                    MessageBox.Show("Save error! \n You need to correctly set the path to save the file. Try again");
                    SaveAs();
                    try
                    {
                        _workbook.SaveAs(path);
                        break;
                    }
                    catch
                    {
                    }
                }
                while (true);
                using (StreamWriter stream = new StreamWriter("path.txt")) { stream.Write(path); }
                MessageBox.Show("Results was saved successfully!");
            }
            while (true)
                try
                {
                    _workbook.Close();
                    _excel.Quit();
                }
                catch
                {
                    _excel.Quit();
                    break;
                }
            return path;
        }
        public void Dispose()
        {
           
        }
    }
}
