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

namespace MailerGUI
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string FromAdress = MailAdress.Text;
            string MailPass = MailPassword.Password;
            string sub = Subject.Text;
            string des = ToAdress.Text;
            string bod = Body.Text;

            var smtp = new System.Net.Mail.SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,

                Credentials = new System.Net.NetworkCredential(FromAdress, MailPass),

                EnableSsl = true
            };

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open("filename");

            workbook.Close();
            excel.Quit();

            try
            {
                var Msg = new System.Net.Mail.MailMessage(FromAdress, des, sub, bod);
                smtp.Send(Msg);
            }
            catch(Exception)
            {
                MessageBox.Show("Exception occurred!");
            }
            finally
            {
                MessageBox.Show("Success!");
            }

            
        }
    }
}
