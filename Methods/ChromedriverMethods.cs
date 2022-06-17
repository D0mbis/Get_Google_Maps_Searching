using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using System.Windows;

namespace Selenium
{
    internal class ChromedriverMethods : IDisposable
    {

        public ChromedriverMethods(string link)
        {
            this.link = link;
        }
        private string link;
        Dictionary<string, List<string>> DictionaryOfResults = new Dictionary<string, List<string>>();

        IWebDriver LaunchChromDriver()
        {
            try
            {
                IWebDriver driver = new ChromeDriver { Url = link };
                return driver;
            }
            catch
            {
                MessageBox.Show("Error launch chromedriver! \n Maybe chromedriver version was chenged.");
                return null;
            }
        }
        public void GetListOfWebElements(int scrollingDelay)
        {
            var driver = LaunchChromDriver();
            Thread.Sleep(3000);
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            var buttonNextPage = driver.FindElement(By.CssSelector("#ppdPk-Ej1Yeb-LgbsSe-tJiF1e"));
            while (true)
            {
                try
                {
                    var listOfWebElements = driver.FindElements(By.XPath(".//*[@aria-label and @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd' or @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t']/div[position()>2]"));
                    int counter = 2;
                    while (counter != 0)
                    {
                        jsExecutor.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, 5000);");
                        Thread.Sleep(scrollingDelay);
                        var ListOfOnePageAfterScrolling = driver.FindElements(By.XPath(".//*[@aria-label and @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd' or @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t']/div"));
                        if (ListOfOnePageAfterScrolling.Count > listOfWebElements.Count)
                        {
                            listOfWebElements = ListOfOnePageAfterScrolling;
                            continue;
                        }
                        else if (ListOfOnePageAfterScrolling.Count <= listOfWebElements.Count)
                        {
                            counter--;
                            continue;
                        }
                    }
                    PutInDictionary(listOfWebElements);
                    buttonNextPage.Click();
                    Thread.Sleep(3000);
                }
                catch
                {
                    break;
                }
            }
            try
            {
                driver.Quit();
            }
            catch
            {
                MessageBox.Show("Error! \n Chromedriver did`t finished work successfully.");
            }
        }
        Dictionary<string, List<string>> PutInDictionary(ReadOnlyCollection<IWebElement> listOfWebElements)
        {
            foreach (var item in listOfWebElements)
            {
                try
                {
                    var nameOneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                    if (nameOneRow != null)
                    {
                        var someContent = item.FindElements(By.XPath(".//*[@class='W4Efsd' and @jsinstance]//span"));
                        //var hyphen = item.FindElement(By.XPath(".//*[@class='W4Efsd' and @jsinstance]//span[@aria-hidden]")).ToString();
                        List<string> contentOneRow = new List<string>();
                        if (someContent != null)
                        {
                            foreach (var item2 in someContent)
                                if (item2.Text != null)
                                {
                                    if (!contentOneRow.Contains($"{item2.Text}") && !item2.Text.ToString().Contains($"·") && !string.IsNullOrEmpty(item2.Text))
                                    {
                                        contentOneRow.Add(item2.Text.ToString());
                                    }
                                    continue;
                                }
                        }
                        DictionaryOfResults[nameOneRow.Text] = contentOneRow;
                    }
                }
                catch
                {
                    continue;
                }
            }
            return DictionaryOfResults;
        }
        public string SaveDictionaryInExcel()
        {
            try
            {
                using (ExcelMethods excelMethods = new ExcelMethods())
                {
                    excelMethods.Open();
                    int rowNumber = 1, columnNumber;
                    foreach (var item in DictionaryOfResults)
                    {
                        columnNumber = 1;
                        excelMethods.ToExcel(rowNumber, columnNumber, item.Key);
                        foreach (var item2 in item.Value)
                        {
                            columnNumber++;
                            excelMethods.ToExcel(rowNumber, columnNumber, item2);
                        }
                        rowNumber++;
                    }
                    //pathValue.Text = excelMethods.Save();
                    MessageBox.Show($"The program successfully finished work. \n Results were found: {DictionaryOfResults.Count}");
                    return excelMethods.Save();
                }
            }
            catch
            {
                MessageBox.Show($"Error! \n Put dictionary in Excel does`t complete successfully.");
                return null;
            }
        }

        public void Dispose()
        {
        }
    }



    /* PhantomJSDriverService driverService = PhantomJSDriverService.CreateDefaultService();
     driverService.HideCommandPromptWindow = true;
 IWebDriver PJS = new PhantomJSDriver(driverService);*/  //HIDE console Window FJS

}
