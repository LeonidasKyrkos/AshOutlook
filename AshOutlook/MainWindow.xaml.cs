using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace AshOutlook
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

        public void browsebtn_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            filepathtxt.Text = dialog.SelectedPath;

        }

        private void recipienttxt_GotFocus(object sender, RoutedEventArgs e)
        {
            recipienttxt.Text = null;
        }

        private void send_Click(object sender, RoutedEventArgs e)
        {            
            
            string folder = filepathtxt.Text;

            if (folder != "")
            {

                string[] filepaths = Directory.GetFiles(folder);

                string email = recipienttxt.Text;

                if (email != "" && email != "Enter recipient email address")
                {

                    foreach (string name in filepaths)
                    {

                        Outlook.Application oApp = new Outlook.Application();
                        Outlook.MailItem nMail = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                        nMail.Subject = "test send";
                        nMail.To = email;
                        nMail.Attachments.Add(name);
                        nMail.Send();
                        oApp.Application.Quit();

                    }
                }

                else
                {
                    System.Windows.MessageBox.Show("Whoops, you forgot the email address");
                }

            }
            else
            {
                System.Windows.MessageBox.Show("Whoops, you forgot to choose the folder");
            }            
         }            
     }
}



            

            
