using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
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

        private void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            if (DateFrome.SelectedDate == null || DateTo.SelectedDate == null || string.IsNullOrEmpty(Value.Text) )
            {
                MessageBox.Show("Please set all cells and try again.");
                return;
            }
            var dateFrom = (DateTime)DateFrome.SelectedDate;
            var dateTo = (DateTime)DateTo.SelectedDate;
            var value = Value.Text;
            IWebDriver driver = new ChromeDriver();
            driver.Url = @"https://www.google.com/maps/search/%D1%80%D0%B5%D1%81%D1%82%D0%BE%D1%80%D0%B0%D0%BD+%D0%B3%D1%80%D1%83%D0%B7%D0%B8%D1%8F/@47.2232698,39.697817,12z/data=!3m1!4b1";
        }
    }
}
