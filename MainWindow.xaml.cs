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
            if (string.IsNullOrEmpty(Value.Text))
            {
                MessageBox.Show("Будь уважніше, всавив якусь херню! Спробуй спочатку");
                return;
            }
            var value = Value.Text;
            IWebDriver driver = new ChromeDriver {Url = @value};
        }
    }
}
