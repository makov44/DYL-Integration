using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
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
using mshtml;
using System.Threading;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;
using Clipboard = System.Windows.Clipboard;
using MessageBox = System.Windows.MessageBox;
using TextDataFormat = System.Windows.TextDataFormat;
using WebBrowser = System.Windows.Forms.WebBrowser;
using System.ComponentModel;
using log4net;

namespace WpfApplication1
{       
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static readonly DependencyProperty UrlProperty =
            DependencyProperty.Register("Url", typeof(string), typeof(MainWindow), new UIPropertyMetadata(string.Empty));

        public string Url
        {
            get { return (string)this.GetValue(UrlProperty); }
            set { this.SetValue(UrlProperty, value); }
        }


        private HtmlDocument document;
        private bool newEmailEventHandeled = false;
        private string txtUrl;
        private string startEmailPage = @"https://webmail.allstate.com";

        public MainWindow()
        {
            InitializeComponent();
            tbUrl.DataContext = this;
            Browser.DocumentCompleted += BrowserOnDocumentCompleted;

            var webBrowser = Browser.ActiveXInstance as SHDocVw.WebBrowser;
            if (webBrowser != null)
            {    
                webBrowser.NewWindow3 += MainWindow_NewWindow;
                webBrowser.Silent = true;
            }

            Url = startEmailPage;


            Browser.Navigate(Url);

        }

        private void BrowserOnDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs webBrowserDocumentCompletedEventArgs)
        {
            document = Browser.Document;
            Url = document?.Url?.ToString();
        }

        private void MainWindow_NewWindow(ref object ppDisp, ref bool Cancel, uint dwFlags, string bstrUrlContext, string bstrUrl)
        {
            Cancel = true;
            Browser.Navigate(bstrUrl);
        }

        private void New_Click(object sender, RoutedEventArgs e)
        {
            var element =  document.GetElementById("newmsgc");
            if (element != null)
            {
                Browser.DocumentCompleted += NewEmail_Click;
                element.InvokeMember("click");
            }

        }

        private void NewEmail_Click(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
           NewEmailHandler();
           Browser.DocumentCompleted -= NewEmail_Click;
        }

        private void NewEmailHandler()
        {
            Thread.Sleep(500);
            var divTo = document.GetElementById("divTo");

            if (divTo == null)
            {
                MessageBox.Show("divTo is NULL");
                Log.Warn("DivTo is NULL");
            }

            if (divTo != null)
                divTo.InnerText = "makov.sergey@gmail.com";


            var txtSubj = document.GetElementById("txtSubj");

            if (txtSubj == null)
            {
                MessageBox.Show("txtSubj is NULL");
                Log.Warn("txtSubj is NULL");
            }

            if (txtSubj != null)
            {
                txtSubj?.Focus();
            }

            Thread.Sleep(200);
            SendKeys.SendWait("subject");
            Thread.Sleep(200);

            var win = document?.Window?.Frames?[0];
            if (win == null)
                return;
            
                var innerdoc = win.Document;
                var frame = innerdoc?.GetElementById("ifBdy");
                frame?.SetAttribute("src", "");


        var frameInnerHtml = @"<html>
	<head>
		<title>Insurance Quote Requested</title>
	</head>
	<body style=""background-color: #ffffff; margin: 0px;padding:0px;font-size:1.1em;font-face:arial, sans-serif;font-weight:normal;"">
		<div style=""clear: both; display: block; width: 600px;background-color: #f4f4f4;margin:0px;padding:0px;"">
			<div style=""padding:0px;margin:0px;padding-top: 45px; margin-left:50px;margin-right:50px;"">
                        <h3 class=""textshadow"" style=""margin-bottom:10; margin-left: 0px; padding:0; font-size: 1.6em; font-weight: bold; color:#424242;"">
			    Insurance Quote Requested
			</h3>
                        <table cellspacing=""0"" cellpadding=""0"" border=""0"" align=""left"" width=""110"">
			    <tr>
				<td valign=""top"" style=""padding-right: 10px; padding-top: 5px;padding-bottom:10px;"">
				    <img width=""100"" height=""100"" src=""cid:test-photo.jpg"" alt="""" align=""left"" style=""border-width: 3px; border-style: solid; border-color: #ffffff;"" />
				</td>
			    </tr>
			</table>
                            <p>
                            Your request for an insurance quote has been received by our agency.<br/><br/>Your quote is currently being prepared. We will contact you shortly with your quote.<br/><br/>To receive your quote now please call us directly at the phone number below.<br/><br/>We appreciate your business and we look forward to serving your insurance needs.<br/>
                            </p>
			</div>
                        <div style=""width: 600px;"">
                            <div style=""width: 600px;background-color: #f9f9f9; border-top: 4px solid #ffffff;margin:0px;padding:0px;padding-top:15px;padding-bottom:20px;"">
                                    <table border=""0"" width=""600"" cellpadding=""0"" cellspacing=""0"">
                                            <tr valign=""top"">
                                                    <td width=""150"" valign=""top"">
                                                        <img width="""" height="""" src=""cid:_agent_photo.png"" style=""margin-top: 6px;margin-left:50px;""/><br/>   
                                                    </td>
                                                    <td width=""12"">&nbsp;</td>
                                                    <td width=""200"">
														<div style=""font-size:1em;color:#424242;padding-top:2px;"">
<strong></strong>
</div>

	
	


<div style=""font-size:.8em; color:#424242;margin-top: 0px;"">
		<strong><em>Email: <a style=""font-size:.8em;color:#424242; text-decoration: none;"" href=""mailto:""></a></em></strong>
</div>
<div style=""font-size:.8em; color:#424242;margin-top: 0px;"">
		<strong><em>Phone: (781) 864-5992</em></strong>
</div>

<div style=""margin-top: 0px;line-height:.9em;"">
	<span style=""color:#868686;font-size:.8em;font-weight:normal;"">
		<br/>
		Cambridge, MA 02141
	</span>
</div>
<div style=""margin-top: 2px;"">
	
		
		
		
		
	
</div>


                                                    </td>
                                                    <td valign=""top"">
                                                            <table cellspacing=""0"" cellpadding=""0"" style=""float:right;margin-right:35px;margin-top:15px;"">
                                                                    <tr>
                                                                        <td><a href=""http://dial-mail.com/email/8405277306/callnow""><img border=""0"" height=""43"" width=""138"" src=""cid:call-now-email.png"" /></a></td>
                                                                    </tr>
                                                                    <tr>
                                                                        <td align=""left"" style=""padding-top:13px;padding-left:5px;"">
                                                                            <span style=""color:#979797;font-size:1em;"">Number On File</span><br/>
                                                                            <span style=""color:#000000;font-size:1.3em;font-weight:bold;"">(614) 333-2217</span>    
                                                                        </td>
                                                                    </tr>
                                                            </table>
                                                    </td>
                                            </td>
                                    </table>
                                </div>
			
		</div>
            </div>
		<img border=""0"" height=""1"" src=""http://dial-mail.com/email/8405277306/pixel.gif"" width=""1""/>
	</body>
</html>";

            frameInnerHtml = frameInnerHtml.Replace("\"\"", "\"");
            //     Clipboard.SetText(body, TextDataFormat.Html);
            //     new InputSimulator().Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);

         //   if (frame != null)
              //  frame.InnerText = frameInnerHtml;
        }

        private void Send_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sendButton = document.GetElementById("send");
                sendButton?.Focus();
                sendButton?.InvokeMember("click");
                Thread.Sleep(1000);
                Browser.Navigate(startEmailPage);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                MessageBox.Show("Send_Click Error");
            }
           
        }

        private void Refresh_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(Url))
                    return;
                else if (Url == Browser.Url?.ToString())
                    Browser.Refresh();
                else
                    Browser.Navigate(Url);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                MessageBox.Show("Refresh_OnClick Error");
            }
           
        }
    }

}
