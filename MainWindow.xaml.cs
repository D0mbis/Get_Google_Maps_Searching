using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
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
        }

        private async void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            bool avalible = false;
            int scrollingDelay = 0, value = 2000, value1 = 1500, value2 = 1000;

            if (string.IsNullOrEmpty(Value.Text))
            {
                MessageBox.Show("Відсутнє посилання для пошуку! Спробуй спочатку");
                return;
            }
            else if (!Value.Text.Contains("https://www.google.com/maps/search/"))
            {
                MessageBox.Show("Посилання має бути з Google maps! Спробуй спочатку");
                return;
            }
            else if (Value.Text.Contains("https://www.google.com/maps/search/"))
            {

                if (ScrollingDelay.IsChecked == true)
                {
                    scrollingDelay = value;
                    avalible = true;
                }

                else if (ScrollingDelay1.IsChecked == true)
                {
                    scrollingDelay = value1;
                    avalible = true;
                }
                else if (ScrollingDelay2.IsChecked == true)
                {
                    scrollingDelay = value2;
                    avalible = true;
                }
                else
                {
                    MessageBox.Show("Не обрана швидкість прокрутки");
                    return;
                }
            }

            if (avalible)
            {
                IWebDriver driver = new ChromeDriver { Url = Value.Text };
                await Task.Delay(3500);
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                var buttonNextPage = driver.FindElement(By.CssSelector("#ppdPk-Ej1Yeb-LgbsSe-tJiF1e"));
                string list = null;
                int counterOfMarks = 0;

                while (true)
                {
                    await Task.Delay(2000);
                    var ListOfOnePage = driver.FindElements(By.XPath(".//*[@aria-label and @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd' or @class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t']/div[position()>2]"));
                    int counter = 2;

                    // scrolling to down page
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

                    // save data from ListOfOnePage
                    foreach (var item in ListOfOnePage)
                    {
                        try
                        {
                            var oneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                            if (oneRow != null)
                            {
                                list += oneRow.Text + " | ";
                                counterOfMarks++;
                            }
                        }
                        catch (System.Exception)
                        {

                            continue;
                        }
                    }

                    // output data from ListOfOnePage
                    try
                    {
                        list += "\n";
                        buttonNextPage.Click();
                    }
                    catch
                    {
                        MessageBox.Show($"Программа завершила работу.\n {list} \n Всего найдено результатов: {counterOfMarks}");
                        driver.Quit();
                        break;
                    }
                }
            }
        }
    }
}