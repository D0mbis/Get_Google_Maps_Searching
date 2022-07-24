using System.Diagnostics;
using System.Windows;

namespace Selenium
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //using (StreamReader stream = new StreamReader("path.txt")) { pathValue.Text = stream.ReadToEnd(); if (!File.Exists(pathValue.Text)) pathValue.Text = null; }
            pathValue.Text = System.Configuration.ConfigurationManager.AppSettings["path"];
        }

        private void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            // checking user input
            using (CheckUserInput userInput = new CheckUserInput(linkValue.Text, ScrollingDelay.IsChecked,
                                                                ScrollingDelay1.IsChecked, ScrollingDelay2.IsChecked))
            {
                if (userInput.Available)
                    using (ChromedriverMethods session = new ChromedriverMethods())
                    {
                        if (session.GetListOfWebElements(userInput.ScrollingDelay, linkValue.Text))
                        {
                            pathValue.Text = session.SaveResultsInExcel();
                        }

                    }
                userInput.AppConfigInfo(pathValue.Text);
            }
            //using (StreamWriter write = new StreamWriter("path.txt")) { write.Write(pathValue.Text); }
        }

        private void OpenDialoSaveAsgBtn(object sender, RoutedEventArgs e)
        {
            using (ExcelMethods excelMethods = new ExcelMethods()) { pathValue.Text = excelMethods.SaveAs(); }
            using (CheckUserInput check = new CheckUserInput()) { check.AppConfigInfo(pathValue.Text); }
            Process.Start(Application.ResourceAssembly.Location);
            Application.Current.Shutdown();

        }
    }
}