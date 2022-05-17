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
        }

        private async void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            bool avalible = false;
            int scrollingDelay = 0, value = 2000, value1 = 1500, value2 = 1000;
            // checking true user input
            if (string.IsNullOrEmpty(Value.Text))
            {
                MessageBox.Show("Відсутнє посилання для пошуку! Спробуй спочатку");
                return;
            }
            else if (!Value.Text.Contains("www.google"))
            {
                MessageBox.Show("Посилання має бути з Google maps! Спробуй спочатку");
                return;
            }
            else if (Value.Text.Contains("www.google"))
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
                    MessageBox.Show("Не обрана швидкість пошуку");
                    return;
                }
            }

            if (avalible)
            {
                IWebDriver driver = new ChromeDriver { Url = Value.Text };
                await Task.Delay(3500);
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                var buttonNextPage = driver.FindElement(By.CssSelector("#ppdPk-Ej1Yeb-LgbsSe-tJiF1e"));
                Dictionary<string, List<string>> list = new Dictionary<string, List<string>>();
                int counterOfResults = 0;
                string path = @"D:\\Programming_study\\Selenium\\Result\\list.xls";
                File.Delete(path);
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
                            var nameOneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                            if (nameOneRow != null)
                            {
                                var someContent = item.FindElements(By.XPath(".//*[@class='W4Efsd' and @jsinstance]"));
                                List<string> contentOneRow = new List<string>();
                                if (someContent != null)
                                {
                                    foreach (var item2 in someContent)
                                        if (item2.Text != null)
                                        {
                                            contentOneRow.Add(item2.Text.ToString());
                                        }
                                }
                                list[nameOneRow.Text] = contentOneRow;
                                //list.Add(nameOneRow.Text, contentOneRow);
                                counterOfResults++;
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
                        using (StreamWriter stream = new StreamWriter(path, true))
                        {
                            foreach (var item in list)
                            {
                                stream.Write(item.Key);
                                foreach (var item2 in item.Value)
                                {
                                    stream.Write(item2);
                                }
                                stream.WriteLine();
                            }
                        }
                        /* foreach (var item in list)
                         {
                             System.IO.File.AppendAllText("D:\\Programming_study\\Selenium\\Result\\list.xls", item.Key + item.Value + "\n"); //Либо найти как открывать файл изменять, сохранять и закрывать либо попробовать dictionary
                         }*/

                        driver.Quit();
                        break;
                        //buttonNextPage.Click();
                    }
                    catch
                    {
                        MessageBox.Show($"Программа завершила работу.\n {list} \n Всего найдено результатов: {counterOfResults}");
                        //MessageBox.Show($"Программа завершила работу.\n {list} \n Всего найдено результатов: {counterOfMarks}");

                        driver.Quit();
                        break;
                    }
                }
            }
        }
    }
}