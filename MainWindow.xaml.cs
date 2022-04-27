using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
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
            //var value = Value.Text;
            else if (Value.Text == "1")
            {
                IWebDriver driver = new ChromeDriver { Url = @"https://www.google.com/maps/search/%D1%81%D1%82%D1%83%D0%B4%D0%B8%D1%8F+%D0%B4%D0%B8%D0%B7%D0%B0%D0%B9%D0%BD%D0%B0+%D0%B3%D1%80%D1%83%D0%B7%D0%B8%D1%8F/@41.7754706,38.7761714,6z/data=!3m1!4b1" };

                await Task.Delay(1000);
                var buttonNextPage = driver.FindElement(By.XPath("//button[@id='ppdPk-Ej1Yeb-LgbsSe-tJiF1e'][contains(.,'disabled')]"));
                var buttonNextPage2 = driver.FindElement(By.XPath("//button[@id='ppdPk-Ej1Yeb-LgbsSe-tJiF1e']"));
                while (buttonNextPage == null)
                {
                        await Task.Delay(3000);
                        buttonNextPage.Click();
                }
                MessageBox.Show("work is done!");
            }
        }
    }
}

//a[contains(., 'disabled')]
