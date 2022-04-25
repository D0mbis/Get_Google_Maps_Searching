using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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
        }
    }
}
