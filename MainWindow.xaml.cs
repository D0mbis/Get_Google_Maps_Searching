using System.IO;
using System.Windows;

namespace Selenium
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            using (StreamReader stream = new StreamReader("path.txt")) { pathValue.Text = stream.ReadToEnd(); }
        }

        private void StartBrn_Click(object sender, RoutedEventArgs e)
        {
            // checking true user input
            using (OtherMethods otherMethods = new OtherMethods())
            {
                bool avalible = otherMethods.CheckingUserInput(linkValue.Text) &&
                otherMethods.CheckingRadiobutton(ScrollingDelay.IsChecked, ScrollingDelay1.IsChecked, ScrollingDelay2.IsChecked);
                if (avalible)
                    using (ChromedriverMethods session = new ChromedriverMethods(linkValue.Text))
                    {
                        session.GetListOfWebElements(otherMethods.ScrollingDelay);
                        pathValue.Text = session.SaveDictionaryInExcel();
                    }
            }
        }

        private void OpenDialoSaveAsgBtn(object sender, RoutedEventArgs e)
        {
            using (ExcelMethods excelMethods = new ExcelMethods()) { pathValue.Text = excelMethods.SaveAs(); }
        }
    }
}

/* 
  +  1. Correctly seve file (messageBox over window
     2. Open every item of ListOfOnePage
     3. Save links and telephone numbers
     4. Go on every link and close (frome excel file or from listOfResult)
     5. Search contacts and save to Excel
     6. Hide program work
     7. Fix bugs(comments)
*/
