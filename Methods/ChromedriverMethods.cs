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
        readonly DateTime timeStart = DateTime.Now;
        IWebDriver Driver { get; set; }
        IJavaScriptExecutor JsExecutor { get; set; }
        ReadOnlyCollection<IWebElement> ListOfWebElements { get; set; }
        Dictionary<string, List<string>> DictionaryOfResults = new Dictionary<string, List<string>>();

        void LaunchChromDriver(string link)
        {
            try
            {
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                ChromeOptions option = new ChromeOptions();
                option.AddArgument("--window-position=-32000,-32000");
                //option.AddArgument("--headless");
                Driver = new ChromeDriver(chromeDriverService, option) { Url = link };
                //Driver.Manage().Window.Maximize();
                JsExecutor = (IJavaScriptExecutor)Driver;
            }
            catch
            {
                MessageBox.Show("Error launch chromedriver! \n Maybe chromedriver version was chenged.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        bool WaitPanelInfo(IWebElement item, By locator)
        {
            int delay = 100;
            int counter = 20;
            while (counter != 0)
            {
                item.Click();
                Thread.Sleep(delay);
                try
                {
                    var element = Driver.FindElement(locator);
                    if (element.Displayed && element.Enabled)
                    {
                        var nameOneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                        var check = item.FindElement(By.XPath("//*[@class='DUwDvf fontHeadlineLarge']//span[text()]"));
                        if (nameOneRow.Text == check.Text)
                            return true;
                        else
                        {
                            counter--;
                            delay += 200;
                        }
                    }
                }
                catch
                {
                    delay += 150;
                    counter--;
                    continue;
                }
            }
            return false;
        }
        ReadOnlyCollection<IWebElement> WaitListOfElements(By locator)
        {
            while (true)
            {
                Thread.Sleep(200);
                try
                {
                    return Driver.FindElements(locator); // download packege chromedriver
                }
                catch
                {
                    continue;
                }
            }
        }
        List<string> WaitForCopyButtonEnabled(IWebElement _item)
        {
            string part1 = "//*[@class='RcCsl fVHpi w4vB1d NOE9ve M0S7ae AG25L']/*[contains(@data-item-id,",
                   part2 = ")]/following-sibling::*//img[contains(@src,'content')]/parent::span/parent::button";
            string[] xPath = { $"{part1}'phone'{part2}", $"{part1}'address'{part2}", $"{part1}'authority'{part2}" };
            List<string> contentOneRow = new List<string>();
            try
            {
                var typeOfActivity = _item.FindElement(By.XPath("//div[@class='fontBodyMedium']//button[text()]"));
                contentOneRow.Add(typeOfActivity.Text.ToString());
            }
            catch
            {
                contentOneRow.Add(" ");
            }
            foreach (var item in xPath)
            {
                try
                {
                    var element = _item.FindElement(By.XPath(item));
                    int _switch = 1, delay = 100;
                    while (_switch < 20)
                        try
                        {
                            Thread.Sleep(delay);
                            element.Click();
                            contentOneRow.Add(Clipboard.GetText().ToString());
                            Clipboard.Clear();
                            break;
                        }
                        catch
                        {
                            try
                            {
                                string path = "document.querySelector('.bJzME.Hu9e2e.tTVLSc *[class*=dS8AEf]')";
                                if (_switch % 2 == 0)
                                {
                                    JsExecutor.ExecuteScript($"{path}.scrollBy(0, 300);");
                                    _switch++;
                                    delay += 200;
                                    continue;
                                }
                                else
                                {
                                    JsExecutor.ExecuteScript($"{path}.scrollBy(0, -300);");
                                    _switch++;
                                    delay += 200;
                                }
                            }
                            catch { }
                            continue;
                        }
                }
                catch
                {
                    contentOneRow.Add("");
                    continue;
                }
            }
            return contentOneRow;
        }

        public bool GetListOfWebElements(string link, int scrollingDelay)
        {
            LaunchChromDriver(link);
            string path = "//*[@aria-label and contains(@class,'m6QErb DxyBCb kA9KIf')]//*[contains(@class,'Nv2PK ')]";
            string path2 = "document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd')";
            while (true)
            {
                try
                {
                    ListOfWebElements = WaitListOfElements(By.XPath(path));
                    int counter = 20;
                    while (counter != 0)
                    {
                        JsExecutor.ExecuteScript($"{path2}.scrollBy(0, 5000);");
                        if (scrollingDelay <= 0) { scrollingDelay = 2000; }
                        Thread.Sleep(scrollingDelay);
                        var ListAfterScrolling = Driver.FindElements(By.XPath(path));

                        if (ListAfterScrolling.Count > ListOfWebElements.Count)
                        {
                            ListOfWebElements = ListAfterScrolling;
                        }
                        else if (ListAfterScrolling.Count <= ListOfWebElements.Count)
                        {
                            JsExecutor.ExecuteScript($"{path2}.scrollBy(0, -5000);");
                            counter--;
                        }
                        try
                        {
                            var endOfPage = Driver.FindElement(By.CssSelector(".PbZDve"));
                            break;
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    // next page button click 
                    try
                    {
                        Driver.FindElement(By.CssSelector("*[id*=eY4Fjd]")).Click();
                    }
                    catch { break; }
                }
                catch
                {
                    break;
                }
            }
            return PutInDictionary();
        }
        bool PutInDictionary()
        {
            if (ListOfWebElements.Count == 0)
            {
                var result = MessageBox.Show("Information not found, please try another search link.", "Bad link", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                if (result == MessageBoxResult.Yes)
                {
                    Driver.Quit();
                    return false;
                }
            }
            List<string> listOfCopyes = new List<string>();
            foreach (var item in ListOfWebElements)
            {
                int counter = 2;
                while (counter != 0)
                {
                    try
                    {
                        if (WaitPanelInfo(item, By.CssSelector(".bJzME.Hu9e2e.tTVLSc")))
                        {
                            var nameOneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                            if (!DictionaryOfResults.ContainsKey(nameOneRow.Text))
                            {
                                DictionaryOfResults[nameOneRow.Text] = WaitForCopyButtonEnabled(item);
                            }
                            else
                            {
                                string copy = nameOneRow.Text + " ";
                                while (true)
                                {
                                    if (listOfCopyes.Contains(copy))
                                    {
                                        copy += " ";
                                    }
                                    else
                                    {
                                        DictionaryOfResults[copy] = WaitForCopyButtonEnabled(item);
                                        listOfCopyes.Add(copy);
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                        else break;
                    }
                    catch
                    {
                        JsExecutor.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, 200);");
                        counter--;
                        continue;
                    }
                }
            }
            ListOfWebElements = null;
            try
            {
                Driver.Quit();
                return true;
            }
            catch
            {
                MessageBox.Show("Chromedriver did`t finished work successfully", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                return true;
            }
        }
        public string SaveResultsInExcel()
        {
            while (true)
                try
                {
                    using (ExcelMethods excel = new ExcelMethods())
                    {
                        return excel.SaveExcelFileNew(DictionaryOfResults, timeStart);
                    }
                }
                catch
                {
                    MessageBox.Show("Save dictionary in Excel does`t complete successfully.", "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    continue;
                }
        }

        public void Dispose()
        {
        }
    }
}