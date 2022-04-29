using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading.Tasks;
using System.Windows;


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
                IWebDriver driver = new ChromeDriver { Url = @"https://www.google.com/maps/search/park+usa/@48.8069443,-109.7329701,3z/data=!3m1!4b1" };

                await Task.Delay(5000);

                //var ListOfOnePage = driver.FindElement(By.XPath(".//*[contains(@class,'m6QErb DxyBCb kA9KIf dS8AEf ecceSd QjC7t')]")); // hiden class after scroll
                var ListOfOnePage = driver.FindElements(By.XPath(".//*[contains(@aria-label,'Результаты по запросу')]/div[position()>2]"));
                if (ListOfOnePage != null)
                {
                    foreach (var item in ListOfOnePage)
                    {
                        try
                        {
                            var oneRow = item.FindElement(By.XPath(".//div[@class='qBF1Pd fontHeadlineSmall']"));
                            if (oneRow != null)
                            {
                                MessageBox.Show(oneRow.Text);
                            }
                         
                        }
                        catch (System.Exception)
                        {

                            continue;
                        }
                        
                    }
                }


                /*var buttonNextPage = driver.FindElement(By.XPath("//button[@id='ppdPk-Ej1Yeb-LgbsSe-tJiF1e']"));
                bool ok = buttonNextPage.Selected;
                while (ok)
                {
                    await Task.Delay(3500);
                    buttonNextPage.Click();
                }
                driver.Quit();*/
            }
        }
    }
    // jstcache="185"
}//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]
 //div[@aria-label ='185'] it work!