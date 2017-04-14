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
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace WpfApplication1
{       
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private HtmlDocument document;
        private bool newEmailEventHandeled = false;
        private string txtUrl;
        private string startPage = @"https://webmail.hostallapps.com/owa/";
        private string webEmailPage = @"https://webmail.hostallapps.com/owa/";

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            Browser.DocumentCompleted += BrowserOnDocumentCompleted;
            Browser.StatusTextChanged += Browser_StatusTextChanged;
            Browser.Navigated += Browser_Navigated;
            var webBrowser = Browser.ActiveXInstance as SHDocVw.WebBrowser;
            if (webBrowser != null)
            {    
                webBrowser.NewWindow3 += MainWindow_NewWindow;
                webBrowser.Silent = true;
            }

            Browser.Navigate(startPage);

        }

        private void Browser_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            tbAddress.Text = Browser.Url.ToString();
        }

        private void Browser_StatusTextChanged(object sender, EventArgs e)
        {
            tblStatus.Text = Browser.StatusText;
        }

        private void BrowserOnDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs webBrowserDocumentCompletedEventArgs)
        {
            Log.Debug($"Called BrowserOnDocumentCompleted document Url = {Browser.Document?.Url}");
            document = Browser.Document;
            if (document != null && document.Window != null)
                document.Window.Error += Window_Error;
            DisableAlertWindow();
            NewEmailHandler();
        }

        private void Window_Error(object sender, HtmlElementErrorEventArgs e)
        {
            var url = "";
            try
            {
                url = e.Url.ToString();
            }
            catch
            {
                url = document?.Url?.ToString();
            }
            
            e.Handled = true;
            Log.Error($"Document error happened. Error:{e.Description}, lineNumber:{e.LineNumber}, url:{url}");
        }

        private void DisableAlertWindow()
        {
            var scriptId = "PopupWindowsBlocker";
            var script = document?.GetElementById(scriptId);

            if (script != null)
                return;

            HtmlElement head = document?.GetElementsByTagName("head")[0];
            HtmlElement scriptEl = document?.CreateElement("script");
            IHTMLScriptElement element = (IHTMLScriptElement)scriptEl?.DomElement;
            if (scriptEl == null)
                return;

            scriptEl.Id = scriptId;

            if (element!=null)
                element.text = JavaScripts.AlertBlocker + " " + JavaScripts.OnbeforeunloadBlocker;

            head?.AppendChild(scriptEl);
        }

        private void MainWindow_NewWindow(ref object ppDisp, ref bool Cancel, uint dwFlags, string bstrUrlContext, string bstrUrl)
        {
            Log.Debug($"Called MainWindow_NewWindow, Url = {bstrUrl}");
            Cancel = true;
            Browser.Navigate(bstrUrl);
        }

        private void New_Click(object sender, RoutedEventArgs e)
        {
            var element =  document.GetElementById("newmsgc");
            if (element != null)
            {
                element.InvokeMember("click");
            }
        }

        private void NewEmailHandler()
        {
            var frame = document.GetElementById("ifBdy");
            var txtSubj = document.GetElementById("txtSubj");
            var divTo = document.GetElementById("divTo");

            if (txtSubj == null || divTo == null || frame == null)
            {
                Log.Debug($"It is not a New email page. txtSubj={txtSubj}," +
                          $" divTo={divTo}, frame={frame}, page url={document.Url}");
                return;
            }

            if (txtSubj.GetAttribute("value") != "subject")
                txtSubj.SetAttribute("value", "subject");

            if (divTo.InnerText != "makov.sergey@gmail.com")
                divTo.InnerText = "makov.sergey@gmail.com";

            var scriptId = "replaceIFrameContentIId";
            var script = document.GetElementById(scriptId);
            if (script != null)
            {
                document.InvokeScript("replaceIFrameContent");
                return;
            }

            var head = document?.GetElementsByTagName("head")[0];
            var scriptEl = document?.CreateElement("script");
            var element = (IHTMLScriptElement)scriptEl?.DomElement;

            if (scriptEl == null || element == null)
                return;

            element.text = JavaScripts.ReplaceIFrameContent;
            scriptEl.Id = scriptId;
            head?.AppendChild(scriptEl);

            document.InvokeScript("replaceIFrameContent");

        }

        private void Send_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sendButton = document.GetElementById("send");
                sendButton?.Focus();
                sendButton?.InvokeMember("click");
                Thread.Sleep(1000);
                Browser.Navigate(webEmailPage);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
        }

        private void Refresh_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Browser.Navigate(tbAddress.Text);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
           
        }

        private void OnPropertyChanged(String name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private void TbUrl_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Browser.Navigate(tbAddress.Text);
            }
        }

        private void Back_OnClick(object sender, RoutedEventArgs e)
        {
           Browser.GoBack();
        }

        private void Forward_OnClick(object sender, RoutedEventArgs e)
        {
            Browser.GoForward();
        }
    }

}
