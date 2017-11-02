using ClosedXML.Excel;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using MailKit;
using MimeKit;

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
            string FromAdress = MailAdress.Text;
            string MailPass = MailPassword.Password;
            string sub = Subject.Text;
            string adresses = ListFile.Text;
            string bod = Body.Text;

            var client = new MailKit.Net.Smtp.SmtpClient();
            var workbook = new XLWorkbook();
            var toad = new Object();
            var exnum = new Object();

            bool exceptionIsOcurred = false;

            var ProgWin = new ProgressWindow();

            ProgWin.Show();

            try{

                await client.ConnectAsync("smtp.gmail.com", 465, MailKit.Security.SecureSocketOptions.SslOnConnect);
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(FromAdress,MailPass);

                workbook = new XLWorkbook(adresses);
                IXLWorksheet worksheet = workbook.Worksheet(1);
                int last = worksheet.LastRowUsed().RowNumber();

                for (int i = 2; i <= last; i++)
                {
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
                }
            }
            catch(Exception ex)
            {
                ProgWin.Close();
                client.Disconnect(true);
                exceptionIsOcurred = true;
                MessageBox.Show(ex.Message,"エラー");

                System.IO.StreamWriter writer = new System.IO.StreamWriter("./errorlog.txt",true,new System.Text.UTF8Encoding(false));
                writer.WriteLine("ErrorMessage:" + ex.Message);
                writer.WriteLine("StackTrace:" + ex.StackTrace);
            }

            if(exceptionIsOcurred == true)
            {
                MessageBox.Show("必要事項および説明書を確認の上、もう一度お試しください。");
            }
            else
            {
                ProgWin.Close();
                client.Disconnect(true);
                MessageBox.Show("メールは正常に送信されました。");
            }
        }
    }
}
