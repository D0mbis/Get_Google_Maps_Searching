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
                MessageBox.Show("Будь уважніше, всавив якусь херню! Спробуй спочатку");
                return;
            }
            //var value = Value.Text;
            else if (Value.Text == "1")
            {
                IWebDriver driver = new ChromeDriver { Url = @"https://www.google.com/maps/search/%D1%80%D0%B5%D1%81%D1%82%D0%BE%D1%80%D0%B0%D0%BD+%D0%B3%D1%80%D1%83%D0%B7%D0%B8%D1%8F/@47.2232698,39.697817,12z" };

                await Task.Delay(1000);
                var buttonNextPage = driver.FindElement(By.XPath("//button[@id='ppdPk-Ej1Yeb-LgbsSe-tJiF1e']"));
                if (buttonNextPage != null)
                {
                    buttonNextPage.Click();
                }
            }
        }
    }
}
// html / body / div[3] / div[9] / div[9] / div / div / div[1] / div[2] / div / div[1] / div / div / div[2] / div[2] / div / div[1] / div / button[2]
//id = "ppdPk-Ej1Yeb-LgbsSe-tJiF1e"