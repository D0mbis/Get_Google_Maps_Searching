using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Selenium
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            using (StreamReader stream = new StreamReader("path.txt")) { pathValue.Text = stream.ReadToEnd(); }
        }

        private async void StartBrn_Click(object sender, RoutedEventArgs e)
        {

            bool radioBtnIsChecked = false;
            int scrollingDelay = 0;
            // checking true user input
            {
                int value = 2000, value1 = 1500, value2 = 1000;
                if (string.IsNullOrEmpty(linkValue.Text))
                {
                    MessageBox.Show("You need input link to search.");
                    return;
                }
                else if (!linkValue.Text.Contains("www.google"))
                {
                    MessageBox.Show("Link must contain 'www.google.com/maps/' \n Try again.");
                    return;
                }
                else if (linkValue.Text.Contains("www.google"))
                {

                    if (ScrollingDelay.IsChecked == true)
                    {
                        scrollingDelay = value;
                        radioBtnIsChecked = true;
                    }

                    else if (ScrollingDelay1.IsChecked == true)
                    {
                        scrollingDelay = value1;
                        radioBtnIsChecked = true;
                    }
                    else if (ScrollingDelay2.IsChecked == true)
                    {
                        scrollingDelay = value2;
                        radioBtnIsChecked = true;
                    }
                    else
                    {
                        MessageBox.Show("You need set speed searching.");
                        return;
                    }
                }
            }
            // logic of program
            if (radioBtnIsChecked)
            {
                try
                {
                    IWebDriver driver = new ChromeDriver { Url = linkValue.Text };
                    await Task.Delay(3500);
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    var buttonNextPage = driver.FindElement(By.CssSelector("#ppdPk-Ej1Yeb-LgbsSe-tJiF1e"));
                    Dictionary<string, List<string>> listOfResults = new Dictionary<string, List<string>>();
                    int counterOfResults = 0;

                    // added to Dictionary ListOfOnePage
                    while (true)
                    {
                        await Task.Delay(2000);
                        var ListOfOnePage = driver.FindElements(By.XPath(".//*[@aria-label and @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd' or @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t']/div[position()>2]"));
                        try
                        {
                            // scrolling to down page
                            int counter = 2;
                            while (counter != 0)
                            {
                                js.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, 5000);");
                                //await Task.Delay(scrollingDelay);
                                Thread.Sleep(scrollingDelay);
                                var ListOfOnePageAfterScrolling = driver.FindElements(By.XPath(".//*[@aria-label and @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd' or @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t']/div"));
                                if (ListOfOnePageAfterScrolling.Count > ListOfOnePage.Count)
                                {
                                    ListOfOnePage = ListOfOnePageAfterScrolling;
                                    continue;
                                }
                                else if (ListOfOnePageAfterScrolling.Count <= ListOfOnePage.Count)
                                {
                                    counter--;
                                    continue;
                                }
                            }
                            // save data to Dictionary ListOfOnePage
                            foreach (var item in ListOfOnePage)
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
                                        listOfResults[nameOneRow.Text] = contentOneRow;
                                        counterOfResults++;
                                    }
                                }
                                catch (System.Exception)
                                {
                                    continue;
                                }
                            }
                            break;
                           // buttonNextPage.Click();
                        }
                        catch
                        {
                            break;
                        }
                    }

                    // save data from ListOfOnePage to Excel
                    try
                    {
                        using (ExcelMethods excelMethods = new ExcelMethods())
                        {
                            excelMethods.Open();
                            int rowNumber = 1, columnNumber;
                            string data;
                            foreach (var item in listOfResults)
                            {
                                columnNumber = 1;
                                data = item.Key;
                                excelMethods.ToExcel(rowNumber, columnNumber, data);
                                foreach (var item2 in item.Value)
                                {
                                    columnNumber++;
                                    data = item2;
                                    excelMethods.ToExcel(rowNumber, columnNumber, data);
                                }
                                rowNumber++;
                            }
                            pathValue.Text = excelMethods.Save();
                        }
                        driver.Quit();
                        MessageBox.Show($"The program successfully finished work. \n Results were found: {counterOfResults}");
                    }
                    catch
                    {
                        MessageBox.Show($"The program does`t finished work successfully.");
                        driver.Quit();
                    }
                }
                catch
                {
                    MessageBox.Show("Error! Maybe chromedriver version was chenged.");
                }
            }
        }

        private void OpenDialoSaveAsgBtn(object sender, RoutedEventArgs e)
        {
            using (ExcelMethods excelMethods = new ExcelMethods()) { pathValue.Text = excelMethods.SaveNewFile(); }
        }
    }
}

   /* 
       1. Correctly seve file
       2. Open every item of ListOfOnePage
       3. Save links and telephone numbers
       4. Go on every link and close (frome excel file or from listOfResult)
       5. Search contacts and save to Excel
       6. Hide program work
       7. Fix bugs
   */
