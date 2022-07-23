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
        DateTime timeStart = DateTime.Now;
        Dictionary<string, List<string>> DictionaryOfResults = new Dictionary<string, List<string>>();

        IWebDriver LaunchChromDriver(string link)
        {
            try
            {
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                //chromeDriverService.HideCommandPromptWindow = true;
                ChromeOptions option = new ChromeOptions();
                //option.AddArgument("--headless");
                IWebDriver driver = new ChromeDriver(chromeDriverService, option) { Url = link };
                driver.Manage().Window.Maximize();
                //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3); 
                return driver;
            }
            catch
            {
                MessageBox.Show("Error launch chromedriver! \n Maybe chromedriver version was chenged.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }
        void WaitPanelInfo(IWebDriver driver, IWebElement item, By locator)
        {
            int delay = 100, counter = 4;
            while (counter != 0)
            {
                item.Click();
                Thread.Sleep(delay);
                try
                {
                    var element = driver.FindElement(locator);
                    if (element.Displayed && element.Enabled)
                    {
                        var nameOneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                        var check = item.FindElement(By.XPath("//*[@class='DUwDvf fontHeadlineLarge']//span[text()]"));
                        if (nameOneRow.Text == check.Text)
                        {
                            bool avalible = item.FindElement(By.XPath("//*[@class='RcCsl fVHpi w4vB1d NOE9ve M0S7ae AG25L']/parent::div")).Enabled;
                            if (avalible)
                                break;
                            else
                            {
                                try { item.FindElement(By.XPath("//*[contains(@class,'VfPpkd-icon-LgbsSe yHy1rc ')]")).Click(); } catch { }
                                counter--;
                                continue;
                            }
                        }
                    }
                    //else continue;
                }
                catch
                {
                    delay += 700;
                    counter--;
                    continue;
                }
            }
        }
        ReadOnlyCollection<IWebElement> WaitListOfElements(IWebDriver driver, By locator)
        {
            while (true)
            {
                Thread.Sleep(200);
                try
                {
                    return driver.FindElements(locator);
                }
                catch
                {
                    continue;
                }
            }
        }
        List<string> WaitForCopyButtonEnabled(IWebElement _item, IJavaScriptExecutor jsExecutor)
        {
            string part1 = "//*[@class='RcCsl fVHpi w4vB1d NOE9ve M0S7ae AG25L']/*[contains(@data-item-id,",
                   part2 = ")]/following-sibling::*//img[contains(@src,'content')]/parent::span/parent::button";
            string[] xPath ={"//div[@class='fontBodyMedium']//button[text()]", $"{part1}'phone'{part2}",
                                              $"{part1}'address'{part2}", $"{part1}'authority'{part2}"};
            List<string> contentOneRow = new List<string>();
            try
            {
                var typeOfActivity = _item.FindElement(By.XPath(xPath[0]));
                contentOneRow.Add(typeOfActivity.Text.ToString());
            }
            catch
            {
                contentOneRow.Add(" ");
            }
            foreach (var item in xPath)
            {
                if (item != xPath[0])
                {
                    int counter = 3;
                    while (counter != 0)
                    {
                        Thread.Sleep(100);
                        try
                        {
                            var element = _item.FindElement(By.XPath(item));
                            if (element.Enabled)
                            {
                                element.Click();
                                if (Clipboard.ContainsText() == true)
                                {
                                    contentOneRow.Add(Clipboard.GetText().ToString());
                                    Clipboard.Clear();
                                    break;
                                }
                                else contentOneRow.Add(""); break;
                            }
                            break;
                        }
                        catch
                        {
                            if (counter == 2)
                            {
                                try
                                {
                                    jsExecutor.ExecuteScript("document.querySelector('.bJzME.Hu9e2e.tTVLSc *[class*=dS8AEf]').scrollBy(0, 300);");
                                    jsExecutor.ExecuteScript("document.querySelector('.bJzME.Hu9e2e.tTVLSc *[class*=dS8AEf]').scrollBy(0, -300);");
                                }
                                catch { }
                                counter--;
                                continue;
                            }
                            if (counter == 1)
                            {
                                contentOneRow.Add(" ");
                                break;
                            }
                            counter--;
                            continue;
                        }
                    }
                }
                else continue;
            }
            return contentOneRow;
        }

        public void GetListOfWebElements(int scrollingDelay, string link)
        {
            var driver = LaunchChromDriver(link);
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            while (true)
            {
                try
                {
                    var listOfWebElements = WaitListOfElements(driver, By.XPath(".//*[@aria-label and contains(@class,'m6QErb DxyBCb kA9KIf')]//*[contains(@class,'Nv2PK ')]"));
                    //int counter = 5;
                    while (true)
                    {
                        jsExecutor.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, 5000);");
                        Thread.Sleep(scrollingDelay);
                        var ListOfOnePageAfterScrolling = driver.FindElements(By.XPath(".//*[@aria-label and contains(@class,'m6QErb DxyBCb kA9KIf')]//*[contains(@class,'Nv2PK ')]"));

                        if (ListOfOnePageAfterScrolling.Count > listOfWebElements.Count)
                        {
                            listOfWebElements = ListOfOnePageAfterScrolling;
                        }
                        else if (ListOfOnePageAfterScrolling.Count <= listOfWebElements.Count)
                        {
                            jsExecutor.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, -5000);");
                            //Thread.Sleep(scrollingDelay);
                            //counter--;
                        }
                        try
                        {
                            var endOfPage = driver.FindElement(By.CssSelector(".PbZDve"));
                            break;
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    PutInDictionary(listOfWebElements, jsExecutor, driver);
                    // next page button click 
                    try
                    {
                        driver.FindElement(By.CssSelector("*[id*=eY4Fjd]")).Click();
                    }
                    catch { break; }
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
        Dictionary<string, List<string>> PutInDictionary(ReadOnlyCollection<IWebElement> listOfWebElements, IJavaScriptExecutor jsExecutor, IWebDriver driver)
        {
            List<string> listOfCopyes = new List<string>();
            foreach (var item in listOfWebElements)
            {
                int counter = 2;
                while (counter != 0)
                {
                    try
                    {
                        WaitPanelInfo(driver, item, By.CssSelector(".bJzME.Hu9e2e.tTVLSc"));
                        var nameOneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                        if (!DictionaryOfResults.ContainsKey(nameOneRow.Text))
                        {
                            DictionaryOfResults[nameOneRow.Text] = WaitForCopyButtonEnabled(item, jsExecutor);
                        }
                        else
                        {
                            //listOfCopyes.Add(nameOneRow.Text);
                            string copy = nameOneRow.Text + " ";
                            while (true)
                            {
                                if (listOfCopyes.Contains(copy))
                                {
                                    copy += " ";
                                }
                                else
                                {
                                    DictionaryOfResults[copy] = WaitForCopyButtonEnabled(item, jsExecutor);
                                    listOfCopyes.Add(copy);
                                    break;
                                }
                            }
                        }
                        break;
                    }
                    catch
                    {
                        jsExecutor.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, 200);");
                        counter--;
                        continue;
                    }
                }
                if (listOfWebElements.IndexOf(item) == listOfWebElements.Count - 1)
                {
                    ;
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
                    return excel.SaveExcelFileNew(DictionaryOfResults, excel.GetCorrectlyPath(), timeStart);
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
