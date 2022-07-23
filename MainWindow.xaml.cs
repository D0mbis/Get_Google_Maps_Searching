using System.IO;
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
                        session.GetListOfWebElements(userInput.ScrollingDelay, linkValue.Text);
                        pathValue.Text = session.SaveResultsInExcel();
                    }
            }
            //using (StreamWriter write = new StreamWriter("path.txt")) { write.Write(pathValue.Text); }
            using (CheckUserInput check = new CheckUserInput()) { check.AppConfigInfo(pathValue.Text); }
        }

        private void OpenDialoSaveAsgBtn(object sender, RoutedEventArgs e)
        {
            using (ExcelMethods excelMethods = new ExcelMethods()) { pathValue.Text = excelMethods.SaveAs(); }
            using (CheckUserInput check = new CheckUserInput()) { check.AppConfigInfo(pathValue.Text); }
        }
    }
}

/* 
  +  1. Correctly seve file  
  +  2. Hide program work
  +  3. Open every item of ListOfOnePage
  +  4. Save links and telephone numbers  
  +  5. FINALY TEST AND CORRECTLY COMMENTS
  +  6. Get rid of "path.txt"
     7. Go on every link and close (frome excel file or from listOfResult)
     8. Search contacts and save to Excel
*/
