using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using ClosedXML.Excel;
using MailKit.Net.Smtp;
using MailKit.Security;
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
            const string closeMessage = "終了しますか？";
            const string closeCaption = "終了";
            const MessageBoxButton closeButton = MessageBoxButton.YesNo;
            const MessageBoxImage closeImage = MessageBoxImage.Question;

            var closeResult = MessageBox.Show(closeMessage, closeCaption, closeButton, closeImage);

            if (closeResult == MessageBoxResult.Yes)
            {
                Environment.Exit(0);
            }
        }

        private async Task LockUi(Func<Task> act)
        {
            if (act == null)
            {
                throw new ArgumentNullException(nameof(act));
            }

            var topElement = ((UIElement)VisualTreeHelper.GetChild(this, 0));

            Cursor = Cursors.Wait;
            topElement.IsEnabled = false;
            await act();

            topElement.IsEnabled = true;
            Cursor = Cursors.AppStarting;

        }

        private async void Button_Click_async(object sender, RoutedEventArgs e)
        {
            if (e == null)
            {
                throw new ArgumentNullException(nameof(e));
            }

            const string messageboxtext = "メールを送信します、よろしいですか？";
            const string caption = "MailerGUI";
            const MessageBoxButton button = MessageBoxButton.OKCancel;
            const MessageBoxImage image = MessageBoxImage.Information;

            var result = MessageBox.Show(messageboxtext, caption, button, image);

            if (result == MessageBoxResult.OK)
            {
                await LockUi(Mailing);
            }
        }

        public async Task Mailing()
        {
#region Datas from user form.
            var fromAdress = MailAdress.Text;
            var mailPass = MailPassword.Password;
            var sub = Subject.Text;
            var adresses = ListFile.Text;
            var bod = Body.Text;
#endregion
            var client = new SmtpClient();

            var xlWorkbook = new XLWorkbook();
            var toad = new object();
            var exnum = new object();

            var unSentListRow = 0;

            var exceptionIsOccurred = false;
            var exceptionInSendProcess = false;

            var progWin = new ProgressWindow();

            progWin.Show();

            try
            {
                client.Connect("smtp.gmail.com", 465, SecureSocketOptions.SslOnConnect);
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(fromAdress, mailPass);
            }
            catch(Exception ex)
            {
                progWin.Close();
                exceptionIsOccurred = true;
                MessageBox.Show(ex.Message, "エラー");
            }

            var unSentAdressList = new XLWorkbook();
            var unSentAdressSheet = unSentAdressList.AddWorksheet("送信に失敗したアドレス");

            try
            {

                var workbook = new XLWorkbook(adresses);
                var worksheet = workbook.Worksheet(1);
                var last = worksheet.LastRowUsed().RowNumber();

                for (var i = 2; i <= last; i++)
                {
                    ContinueSending:

                    try
                    {
                        if (exceptionIsOccurred)
                        {
                            i++;
                        }

                        var message = bod;

                        var cell = worksheet.Cell(i, 1);
                        var num = worksheet.Cell(i, 2);

                        exnum = num.Value;
                        toad = cell.Value;

                        message = message.Replace("repl", exnum.ToString());

                        var msg = new MimeMessage();
                        msg.From.Add(new MailboxAddress(fromAdress));
                        msg.Subject = sub;
                        msg.To.Add(new MailboxAddress(toad.ToString(), toad.ToString()));
                        var msgbuilder = new BodyBuilder {TextBody = message};

                        msg.Body = msgbuilder.ToMessageBody();

                        await client.SendAsync(msg);

                        //これつけないと他の人のメアドが全部Toに表示される。
                        msg.To.RemoveAt(0);
                    }
                    catch(Exception)
                    {
                        unSentListRow++;

                        exceptionInSendProcess = true;

                        var unSentAdress = unSentAdressSheet.Cell(unSentListRow, 1);
                        unSentAdress.SetValue(toad);

                        var unSentNumber = unSentAdressSheet.Cell(unSentListRow, 2);
                        unSentNumber.SetValue(exnum);

                        //送信中に例外拾っても最後まで送る。
                        goto ContinueSending;
                    }
                }
            }
            catch(Exception excelException)
            {
                exceptionIsOccurred = true;
                MessageBox.Show(excelException.Message,"エラー");
            }

            unSentAdressList.SaveAs("送信失敗リスト.xlsx");

            if(exceptionIsOccurred)
            {
                progWin.Close();
                client.Disconnect(true);
                MessageBox.Show("必要事項および説明書を確認の上、もう一度お試しください。");
            }
            if(exceptionInSendProcess)
            {
                progWin.Close();
                client.Disconnect(true);
                MessageBox.Show("一部のメールが正しく送信されませんでした。「送信失敗リスト.xlsx」を参照してください。");
            }
            else
            {
                progWin.Close();
                client.Disconnect(true);
                MessageBox.Show("メールが正常に送信されました。");
            }
        }
    }
}
