using System;
using System.Configuration;
using System.Windows;

namespace Selenium
{
    internal class CheckUserInput : IDisposable
    {
        public int ScrollingDelay { get; private set; }
        public bool Available { get; private set; }
        public void CheckInput(string text, bool? isChecked1, bool? isChecked2, bool? isChecked3)
        {
            Available = CheckingInputLink(text) &&
            CheckingRadiobutton(isChecked1, isChecked2, isChecked3);
        }
        bool CheckingInputLink(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("You need input link to search.");
                return false;
            }
            else if (!text.Contains("www.google"))
            {
                MessageBox.Show("Link must contain 'www.google.com/maps/' \nTry again.");
                return false;
            }
            else if (text.Contains("www.google"))
                return true;
            else return false;
        }

        bool CheckingRadiobutton(bool? isChecked1, bool? isChecked2, bool? isChecked3) 
        {
            int value = 2000, value1 = 1500, value2 = 1000;
            {
                if (isChecked1 == true)
                {
                    ScrollingDelay = value;
                    return true;
                }

                else if (isChecked2 == true)
                {
                    ScrollingDelay = value1;
                    return true;
                }
                else if (isChecked3 == true)
                {
                    ScrollingDelay = value2;
                    return true;
                }
                else
                {
                    MessageBox.Show("You need set speed searching.");
                    return false;
                }
            }
        }
        public void SaveAppConfig(string path)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            if (config.AppSettings.Settings["path"] == null)
            {
                config.AppSettings.Settings.Add("path", path);
            }
            else
            {
                config.AppSettings.Settings["path"].Value = path;
            }
            config.Save();
        }

        public void Dispose()
        {
        }
    }
}