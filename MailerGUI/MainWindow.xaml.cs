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
            string bodyAfter = BodyAfterVals.Text;
            var toad = new Object();
            var exnum = new Object();

            var ProgWin = new ProgressWindow();

            ProgWin.Show();

            var smtp = new System.Net.Mail.SmtpClient
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

                for (int i = 1; i <= last; i++)
                {
                    string depbod;

                    IXLCell cell = worksheet.Cell(i, 1);
                    IXLCell num = worksheet.Cell(i, 2);

                    exnum = num.Value;
                    toad = cell.Value;

                    depbod = bod + exnum + "\n" + bodyAfter;

                    var Msg = new System.Net.Mail.MailMessage(FromAdress, toad.ToString(), sub, depbod);
                    await smtp.SendMailAsync(Msg);
                }
            }
            catch (Exception)
            {
                ProgWin.Close();
                MessageBox.Show("エラーが発生しました。");
            }
            finally
            {
                ProgWin.Close();
                MessageBox.Show("メール送信成功！");
            }
        }
    }
}
