using GetSearching_GM;
using System.Diagnostics;
using System.Threading;
using System.Windows;

namespace Selenium
{
    public partial class MainWindow : Window
    {
        private Progress Progress;
        public MainWindow()
        {
            InitializeComponent();
            pathValue.Text = System.Configuration.ConfigurationManager.AppSettings["path"];
            linkValue.Text = System.Configuration.ConfigurationManager.AppSettings["link"];
        }

        private void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            if (Progress == null)
            {
                using (CheckUserInput userInput = new CheckUserInput())
                {
                    userInput.CheckInput(linkValue.Text, ScrollingDelay.IsChecked,
                                         ScrollingDelay1.IsChecked, ScrollingDelay2.IsChecked);
                    if (userInput.Available)
                        using (ChromedriverMethods session = new ChromedriverMethods())
                        {
                            session.GetListOfWebElements(linkValue.Text, userInput.ScrollingDelay);
                            Thread thread = new Thread(session.PutInDictionary);
                            thread.SetApartmentState(ApartmentState.STA);
                            Progress progress = session;
                            thread.Start();
                            Progress = progress;
                            progress.Show();
                        }
                }
                pathValue.Text = System.Configuration.ConfigurationManager.AppSettings["path"];
            }
            else { Progress.Activate(); }
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