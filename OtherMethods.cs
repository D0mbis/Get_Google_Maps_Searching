using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Selenium
{
    internal class OtherMethods : IDisposable
    {
        public int scrollingDelay { get; set; }

        public bool CheckingUserInput(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("You need input link to search.");
                return false;
            }
            else if (!text.Contains("www.google"))
            {
                MessageBox.Show("Link must contain 'www.google.com/maps/' \n Try again.");
                return false;
            }
            else if (text.Contains("www.google"))
                return true;
            else return false;
        }

        public bool CheckingRadiobutton (bool? isChecked1, bool? isChecked2, bool? isChecked3)
        {
            int value = 2000, value1 = 1500, value2 = 1000;
            {
                if (isChecked1 == true)
                {
                    scrollingDelay = value;
                    return true;
                }

                else if (isChecked2 == true)
                {
                    scrollingDelay = value1;
                    return  true;
                }
                else if (isChecked3 == true)
                {
                    scrollingDelay = value2;
                    return  true;
                }
                else
                {
                    MessageBox.Show("You need set speed searching.");
                    return false;
                }
            }
        }
        public void Dispose()
        {
        }
    }
}
