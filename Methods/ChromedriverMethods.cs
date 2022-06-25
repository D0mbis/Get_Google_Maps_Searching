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
        Dictionary<string, List<string>> DictionaryOfResults = new Dictionary<string, List<string>>();

        IWebDriver LaunchChromDriver(string link)
        {
            try
            {
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                ChromeOptions option = new ChromeOptions();
                //option.AddArgument("--headless");
                IWebDriver driver = new ChromeDriver(chromeDriverService, option) { Url = link };
                return driver;
            }
            catch
            {
                MessageBox.Show("Error launch chromedriver! \n Maybe chromedriver version was chenged.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }
        public void GetListOfWebElements(int scrollingDelay, string link)
        {
            var driver = LaunchChromDriver(link);
            Thread.Sleep(4000);
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            var buttonNextPage = driver.FindElement(By.CssSelector("#ppdPk-Ej1Yeb-LgbsSe-tJiF1e"));
            while (true)
            {
                try
                {
                    var listOfWebElements = driver.FindElements(By.XPath(".//*[@aria-label and @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd' or @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t']/div/div[contains(@class,'Nv2PK ')]"));
                    int counter = 2;
                    while (counter != 0)
                    {
                        jsExecutor.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, 5000);");
                        Thread.Sleep(scrollingDelay);
                        var ListOfOnePageAfterScrolling = driver.FindElements(By.XPath(".//*[@aria-label and @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd' or @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t']/div/div[contains(@class,'Nv2PK ')]"));
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
                    PutInDictionary(listOfWebElements, jsExecutor);
                    buttonNextPage.Click();
                    Thread.Sleep(1000);
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
                MessageBox.Show("Chromedriver did`t finished work successfully", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        Dictionary<string, List<string>> PutInDictionary(ReadOnlyCollection<IWebElement> listOfWebElements, IJavaScriptExecutor jsExecutor)
        {
            foreach (var item in listOfWebElements)
            {
                try
                {
                    var nameOneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                    if (nameOneRow != null)
                    {
                        List<string> contentOneRow = new List<string>();
                        string[] xPath ={"//div[@class='fontBodyMedium']//button[text()]",
                                         "//div[@class='RcCsl fVHpi w4vB1d NOE9ve M0S7ae AG25L']/button[contains(@data-item-id,'phone')]/following-sibling::div//img[contains(@src,'content')]",
                                         "//div[@class='RcCsl fVHpi w4vB1d NOE9ve M0S7ae AG25L']/button[@data-item-id='address']/following-sibling::div//img[contains(@src,'content')]",
                                         "//div[@class='RcCsl fVHpi w4vB1d NOE9ve M0S7ae AG25L']/button[contains(@data-item-id,'authority')]/following-sibling::div//img[contains(@src,'content')]"};
                        item.Click();
                        Thread.Sleep(2500);
                        try
                        {
                            contentOneRow.Add(item.FindElement(By.XPath(xPath[0])).Text.ToString());
                        }
                        catch
                        {
                            contentOneRow.Add(" ");
                        }

                        jsExecutor.ExecuteScript("document.querySelector('div .bJzME.Hu9e2e.tTVLSc div .m6QErb.WNBkOb div.m6QErb.DxyBCb.kA9KIf.dS8AEf').scrollBy(0, 350);");
                        foreach (var item2 in xPath)
                        {
                            if (item2 != xPath[0])
                                try
                                {
                                    item.FindElement(By.XPath(item2)).Click();
                                    Thread.Sleep(500);
                                    if (Clipboard.ContainsText() == true)
                                    {
                                        contentOneRow.Add(Clipboard.GetText().ToString());
                                        Clipboard.Clear();
                                    }
                                    else
                                    {
                                        contentOneRow.Add(" ");
                                        continue;
                                    }
                                }
                                catch
                                {
                                    contentOneRow.Add(" ");
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
        public string SaveResultsInExcel()
        {
            try
            {
                using (ExcelMethods excel = new ExcelMethods())
                {
                    return excel.SaveExcelFileNew(DictionaryOfResults, excel.GetCorrectlyPath());
                }
            }
            catch
            {
                MessageBox.Show("Save dictionary in Excel does`t complete successfully.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public void Dispose()
        {
        }
    }
}