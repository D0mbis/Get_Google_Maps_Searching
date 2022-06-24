using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

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
                else if(string.IsNullOrEmpty(path))
                {
                    var result = MessageBox.Show("Please provide a path to save the file:", "Save results file", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    if(result == MessageBoxResult.OK)
                    {
                        SaveAs();
                        continue;
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    var result = MessageBox.Show($"Do you want save file results in: {path}?", "Worning", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes && path != null)
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
                int rowNumber = 2, columnNumber;
                foreach (var item in dictionaryOfResults)
                {
                    columnNumber = 1;
                    _worksheet.Cells[rowNumber, columnNumber].Value = item.Key;
                    foreach (var item2 in item.Value)
                    {
                        columnNumber++;
                        try
                        {
                            _worksheet.Cells[rowNumber, columnNumber].Value = item2;
                        }
                        catch { }
                    }
                    rowNumber++;
                }
                try
                {
                    _worksheet.Cells.AutoFitColumns();
                    _worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    package.Save();
                    MessageBox.Show($"Results were saved: {dictionaryOfResults.Count}", "Successfully completed", MessageBoxButton.OK, MessageBoxImage.Information);
                    return path;
                }
                catch
                {
                    MessageBox.Show("Folder does not exist, select an existing folder.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    path = SaveAs();
                    package.SaveAs(path);
                    return path;
                }
            }
        }

        public void Dispose()
        {

        }
    }
}
