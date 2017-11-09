using ClosedXML.Excel;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

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
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string CloseMessage = "終了しますか？";
            string CloseCaption = "終了";
            MessageBoxButton CloseButton = MessageBoxButton.YesNo;
            MessageBoxImage CloseImage = MessageBoxImage.Question;

            MessageBoxResult CloseResult = MessageBox.Show(CloseMessage, CloseCaption, CloseButton, CloseImage);

            if (CloseResult == MessageBoxResult.Yes)
            {
                Environment.Exit(0);
            }
        }

        private async Task LockUI(Func<Task> act)
        {
            var topElm = ((UIElement)VisualTreeHelper.GetChild(this, 0));
            var oldEnabled = topElm.IsEnabled;
            var oldCursor = this.Cursor;
            try
            {
                this.Cursor = Cursors.Wait;
                topElm.IsEnabled = false;
                await act();
            }
            finally
            {
                topElm.IsEnabled = oldEnabled;
                this.Cursor = oldCursor;
            }
        }

        private static void ErrorLogWriter(System.Exception ex)
        {
            System.IO.StreamWriter writer = new System.IO.StreamWriter("./errorlog.txt", true, new System.Text.UTF8Encoding(false));
            writer.WriteLine("ErrorMessage:" + ex.Message);
            writer.WriteLine("StackTrace:" + ex.StackTrace);
        }

        private async void Button_Click_async(object sender, RoutedEventArgs e)
        {
            string messageboxtext = "メールを送信します、よろしいですか？";
            string caption = "MailerGUI";
            MessageBoxButton button = MessageBoxButton.OKCancel;
            MessageBoxImage image = MessageBoxImage.Information;

            MessageBoxResult result = MessageBox.Show(messageboxtext, caption, button, image);

            if (result == MessageBoxResult.OK)
            {
                await LockUI(async () => { await Mailing(); });
            }
        }

        public async Task Mailing()
        {
#region Datas from user form.
            string FromAdress = MailAdress.Text;
            string MailPass = MailPassword.Password;
            string sub = Subject.Text;
            string adresses = ListFile.Text;
            string bod = Body.Text;
#endregion
            var client = new MailKit.Net.Smtp.SmtpClient();

            var workbook = new XLWorkbook();
            int last;
            IXLWorksheet worksheet;
            var toad = new Object();
            var exnum = new Object();

            int unSentListRow = 0;

            bool exceptionIsOccurred = false;
            bool exceptionInSendProcess = false;

            var ProgWin = new ProgressWindow();

            ProgWin.Show();

            try
            {
                await client.ConnectAsync("smtp.gmail.com", 465, MailKit.Security.SecureSocketOptions.SslOnConnect);
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(FromAdress, MailPass);
            }
            catch(Exception ex)
            {
                ProgWin.Close();
                exceptionIsOccurred = true;
                MessageBox.Show(ex.Message, "エラー");
            }

            var unSentAdressList = new XLWorkbook();
            var unSentAdressSheet = unSentAdressList.AddWorksheet("送信に失敗したアドレス");

            try
            {

                workbook = new XLWorkbook(adresses);
                worksheet = workbook.Worksheet(1);
                last = worksheet.LastRowUsed().RowNumber();

                for (int i = 2; i <= last; i++)
                {
                    ContinueSending:

                    try
                    {
                        if (exceptionIsOccurred == true)
                        {
                            i++;
                        }

                        string message = bod;

                        IXLCell cell = worksheet.Cell(i, 1);
                        IXLCell num = worksheet.Cell(i, 2);

                        exnum = num.Value;
                        toad = cell.Value;

                        message = message.Replace("repl", exnum.ToString());

                        var Msg = new MimeKit.MimeMessage();
                        Msg.From.Add(new MimeKit.MailboxAddress(FromAdress));
                        Msg.Subject = sub;
                        Msg.To.Add(new MimeKit.MailboxAddress(toad.ToString(), toad.ToString()));
                        var Msgbuilder = new MimeKit.BodyBuilder();

                        Msgbuilder.TextBody = message;
                        Msg.Body = Msgbuilder.ToMessageBody();

                        await client.SendAsync(Msg);

                        Msg.To.RemoveAt(0);
                        exceptionIsOccurred = false;
                    }
                    catch(Exception SendEx)
                    {
                        unSentListRow++;

                        exceptionIsOccurred = true;
                        exceptionInSendProcess = true;

                        IXLCell unSentAdress = unSentAdressSheet.Cell(unSentListRow, 1);
                        unSentAdress.SetValue<object>(toad);

                        IXLCell unSentNumber = unSentAdressSheet.Cell(unSentListRow, 2);
                        unSentNumber.SetValue<object>(exnum);

                        ErrorLogWriter(SendEx);

                        //Continue mailing process when exception occurred.
                        goto ContinueSending;
                    }
                }
            }
            catch(Exception ExcelEx)
            {
                exceptionIsOccurred = true;

                ErrorLogWriter(ExcelEx);
            }

            unSentAdressList.SaveAs("送信失敗リスト.xlsx");

            if(exceptionIsOccurred == true)
            {
                ProgWin.Close();
                client.Disconnect(true);
                MessageBox.Show("必要事項および説明書を確認の上、もう一度お試しください。");
            }
            if(exceptionInSendProcess == true)
            {
                ProgWin.Close();
                client.Disconnect(true);
                MessageBox.Show("一部のメールが正しく送信されませんでした。「送信失敗リスト.xlsx」を参照してください。");
            }
            else
            {
                ProgWin.Close();
                client.Disconnect(true);
                MessageBox.Show("メールが正常に送信されました。");
            }
        }
    }
}
