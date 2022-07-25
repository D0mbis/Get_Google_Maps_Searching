using System.Diagnostics;
using System.Windows;

namespace Selenium
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            pathValue.Text = System.Configuration.ConfigurationManager.AppSettings["path"];
        }

        private void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            // checking user input
            using (CheckUserInput userInput = new CheckUserInput())
            {
                userInput.CheckInput(linkValue.Text, ScrollingDelay.IsChecked,
                                     ScrollingDelay1.IsChecked, ScrollingDelay2.IsChecked);
                if (userInput.Available)
                    using (ChromedriverMethods session = new ChromedriverMethods())
                    {
                        if (session.GetListOfWebElements(userInput.ScrollingDelay, linkValue.Text))
                        {
                            pathValue.Text = session.SaveResultsInExcel();
                        }

                    }
                userInput.SaveAppConfig(pathValue.Text);
            }
        }

        private void OpenDialoSavePathBtn(object sender, RoutedEventArgs e)
        {
            using (ExcelMethods excelMethods = new ExcelMethods()) { pathValue.Text = excelMethods.SaveAs(); }
            using (CheckUserInput check = new CheckUserInput()) { check.SaveAppConfig(pathValue.Text); }
            // reload App from update AppConfig
            Process.Start(Application.ResourceAssembly.Location);
            Application.Current.Shutdown();
        }
    }
}