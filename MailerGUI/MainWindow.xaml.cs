using ClosedXML.Excel;
using System;
using System.Windows;

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
            string adresses = ListFile.Text;
            string bod = Body.Text;
            string bodyAfter = BodyAfterVals.Text;
            var toad = new Object();        //Variable for saving values from ClosedXML.(toad/exnum)
            var exnum = new Object();

            var smtp = new System.Net.Mail.SmtpClient   //Gmail authorization.
            {
                Host = "smtp.gmail.com",
                Port = 587,

                Credentials = new System.Net.NetworkCredential(FromAdress, MailPass),

                EnableSsl = true
            };

            XLWorkbook workbook = new XLWorkbook(adresses);
            IXLWorksheet worksheet = workbook.Worksheet(1);

            int last = worksheet.LastRowUsed().RowNumber();

            try
            {
                var ProgWin = new ProgressWindow();

                for (int i = 1; i <= last; i++)     //Mailing process. Use .xlsx file for resource.
                {
                    string depbod;

                    IXLCell cell = worksheet.Cell(i, 1);
                    IXLCell num = worksheet.Cell(i, 2);

                    exnum = num.Value;
                    toad = cell.Value;

                    depbod = bod + exnum + bodyAfter;

                    var Msg = new System.Net.Mail.MailMessage(FromAdress, toad.ToString(), sub, bod);
                    smtp.Send(Msg);
                }
            }
            catch(Exception)
            {
                MessageBox.Show("エラーが発生しました。");
            }
            finally
            {
                MessageBox.Show("メール送信成功！");
            }

            //try
            //{
            //    var Msg = new System.Net.Mail.MailMessage(FromAdress, toad.ToString(), sub, bod);
            //    //var Msg = new System.Net.Mail.MailMessage(FromAdress, to, sub, bod);
            //    smtp.Send(Msg);
            //}
            //catch(Exception)
            //{
            //    MessageBox.Show("Exception occurred!");
            //}
            //finally
            //{
            //    MessageBox.Show("Success!");
            //}
            //These codes won't be used.(But leave for guarantee)
            
        }
    }
}
