using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

//using Excel = Microsoft.Office.Interop.Excel;

namespace Selenium
{
    internal class ExcelMethods : IDisposable
    {

        private string path;

        public ExcelMethods()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (StreamReader stream = new StreamReader("path.txt")) { path = stream.ReadToEnd(); }

        }
        public string GetCorrectlyPath()
        {
            while (true)
            {
                if (File.Exists(path))
                {
                    try
                    {
                        var result = MessageBox.Show($"File: {path} will be overwrite. \n Are you sure?", "Worning", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (result == MessageBoxResult.Yes)
                        {
                            File.Delete(path);
                            return path;
                        }
                        else
                        {
                            SaveAs();
                            continue;
                        }
                    }
                    catch
                    {
                        continue;
                    }

                }
                else
                {
                    var result = MessageBox.Show($"Do you want save file results in: {path}?", "Worning", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        return path;
                    }
                    else
                    {
                        SaveAs();
                        continue;
                    }
                }
            }
        }

        public string SaveAs()
        {
            while (true)
            {
                try
                {
                    SaveFileDialog saveDialog = new SaveFileDialog
                    {
                        Filter = "Книга Excel 97-2003 (*.xlsx) | *.xlsx"
                    };
                    if (saveDialog.ShowDialog() == true)
                    {
                        path = saveDialog.FileName.ToString();
                    }
                    return path;
                }
                catch
                {
                    continue;
                }
            }
        }

        public string SaveExcelFileNew(Dictionary<string, List<string>> dictionaryOfResults, string path)
        {

            using (ExcelPackage package = new ExcelPackage(path))
            {
                var _worksheet = package.Workbook.Worksheets.Add("Results");
                int rowNumber = 1, columnNumber;
                foreach (var item in dictionaryOfResults)
                {
                    columnNumber = 1;
                    _worksheet.Cells[rowNumber, columnNumber].LoadFromText(item.Key);
                    foreach (var item2 in item.Value)
                    {
                        columnNumber++;
                        try
                        {
                            _worksheet.Cells[rowNumber, columnNumber].LoadFromText(item2);
                        }
                        catch { }
                    }
                    rowNumber++;
                }

                try
                {
                    _worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    _worksheet.Cells.AutoFitColumns();
                    package.Save();
                    MessageBox.Show($"Results were saved: {dictionaryOfResults.Count}", "Successfully completed", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch 
                {
                    MessageBox.Show("Save Excel file does`t complete successfully.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                return path;
            }
        }

        public void Dispose()
        {

        }
    }
}
