using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;

namespace Selenium
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Value.Text))
            {
                MessageBox.Show("Будь уважніше, вставив якусь херню! Спробуй спочатку");
                return;
            }
            else if (Value.Text == "1")
            {
                IWebDriver driver = new ChromeDriver { Url = @"https://www.google.com/maps/search/%D0%BA%D0%B0%D1%84%D0%B5+%D0%B2%D0%BE%D0%BB%D0%B3%D0%BE%D0%B4%D0%BE%D0%BD%D1%81%D0%BA/@47.5162181,42.1873579,13z/data=!3m1!4b1" };
                await Task.Delay(3500);
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                var buttonNextPage = driver.FindElement(By.CssSelector("#ppdPk-Ej1Yeb-LgbsSe-tJiF1e"));
                var buttonNextPageDisable = driver.FindElement(By.CssSelector("#ppdPk-Ej1Yeb-LgbsSe-tJiF1e[disabled]"));
                while (true)
                {
                    await Task.Delay(2000);
                    var ListOfOnePage = driver.FindElements(By.XPath(".//*[contains(@aria-label,'Результаты по запросу')]/div[position()>2]"));
                    int counter = 2;
                    // scrolling to down page
                    while (counter != 0)
                    {
                        js.ExecuteScript("document.querySelector('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollBy(0, 5000);");
                        await Task.Delay(2000);
                        var ListOfOnePageAfterScrolling = driver.FindElements(By.XPath(".//*[contains(@aria-label,'Результаты по запросу')]/div[position()>2]"));
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
                    // output data from ListOfOnePage
                    string list = null;
                    foreach (var item in ListOfOnePage)
                    {
                        
                        try
                        {
                            var oneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                            if (oneRow != null)
                            {
                                list += "\n" + oneRow.Text;   
                            }
                        }
                        catch (System.Exception)
                        {
                            continue;
                        }
                    }
                    MessageBox.Show(list);
                    if (buttonNextPage.Displayed)
                    {
                        ListOfOnePage = null;
                        buttonNextPage.Click();
                    }
                    else if (buttonNextPageDisable )  //  ДОДЕЛАТЬ УСЛОВИЕ ОБНАРУЖЕНИЯ НЕДОСТУПНОСТИ КНОПКИ disabled
                    {
                        driver.Quit();
                        break;
                    }
                }
            }
        }
    }
}

//// WORKING SCROLLING
//document.querySelector('#QA0Szd > div > div > div.w6VYqd > div.bJzME.tTVLSc > div > div.e07Vkf.kA9KIf > div > div > div.m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd > div.m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd').scrollTop = 200;
//$$('[aria-label].m6QErb.DxyBCb.kA9KIf.dS8AEf.ecceSd')