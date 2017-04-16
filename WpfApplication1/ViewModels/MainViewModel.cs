using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using DYL.EmailIntegration.Helpers;
using log4net;
using mshtml;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace DYL.EmailIntegration.ViewModels
{
    public class MainViewModel: ViewModelBase
    {
        public ICommand LogInCommand => new DelegateCommand(param=> { LogIn_OnClick(); });
        public ICommand LogOutCommand => new DelegateCommand(param => { LogOut_OnClick(); });
        public ICommand BackCommand => new DelegateCommand(param => { Back_OnClick();});
        public ICommand ForwardCommand => new DelegateCommand(param => { Forward_OnClick();});
        public ICommand NewCommand => new DelegateCommand(param=> { New_OnClick(); });
        public ICommand SendCommand => new DelegateCommand(param => { Send_OnClick(); });
        public ICommand RefreshCommand => new DelegateCommand(param => { Refresh_OnClick(); });
        public ICommand EnterCommand => new DelegateCommand(param => { TbUrl_OnKeyDown(param?.ToString()); });

        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private WebBrowser _browser;
        private HtmlDocument document;
        private string startPage = @"https://webmail.hostallapps.com/owa/";
        private string webEmailPage = @"https://webmail.hostallapps.com/owa/";

        private string _address;
        public string Address
        {
            get { return _address; }
            set
            {
                _address = value;
                RaisePropertyChangedEvent("Address");
            }
        }

        private string _status;
        public string Status
        {
            get { return _status; }
            set
            {
                _status = value;
                RaisePropertyChangedEvent("Status");
            }
        }

        private Visibility _mainLayoutVisibility;
        public Visibility MainLayoutVisibility
        {
            get { return _mainLayoutVisibility; }
            set
            {
                _mainLayoutVisibility = value;
                RaisePropertyChangedEvent("MainLayoutVisibility");
            }
        }

        private Visibility _logInLayoutVisibility;
        public Visibility LogInLayoutVisibility
        {
            get { return _logInLayoutVisibility; }
            set
            {
                _logInLayoutVisibility = value;
                RaisePropertyChangedEvent("LogInLayoutVisibility");
            }
        }

        public MainViewModel(WebBrowser browser)
        {
            _browser = browser;
            _browser.DocumentCompleted += BrowserOnDocumentCompleted;
            _browser.StatusTextChanged += Browser_StatusTextChanged;
            _browser.Navigated += Browser_Navigated;
            var webBrowser = _browser.ActiveXInstance as SHDocVw.WebBrowser;
            if (webBrowser != null)
            {
                webBrowser.NewWindow3 += MainWindow_NewWindow;
                webBrowser.Silent = true;
            }
            _browser.Navigate(startPage);
            MainLayoutVisibility = Visibility.Collapsed;
            LogInLayoutVisibility = Visibility.Visible;
        }

        #region Browser events
        private void Browser_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            Address = _browser.Url.ToString();
        }

        private void Browser_StatusTextChanged(object sender, EventArgs e)
        {
            Status = _browser.StatusText;
        }

        private void BrowserOnDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs webBrowserDocumentCompletedEventArgs)
        {
            Log.Debug($"Called BrowserOnDocumentCompleted document Url = {_browser.Document?.Url}");
            document = _browser.Document;
            if (document != null && document.Window != null)
                document.Window.Error += Window_Error;
            DisableAlertWindow();
            NewEmailHandler();
        }

        private void MainWindow_NewWindow(ref object ppDisp, ref bool Cancel, uint dwFlags, string bstrUrlContext,
            string bstrUrl)
        {
            Log.Debug($"Called MainWindow_NewWindow, Url = {bstrUrl}");
            Cancel = true;
            if(!string.IsNullOrEmpty(bstrUrl))
                _browser.Navigate(bstrUrl);
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

        private void Send_OnClick()
        {
            try
            {
                var sendButton = document.GetElementById("send");
                sendButton?.Focus();
                sendButton?.InvokeMember("click");
                Thread.Sleep(500);
                _browser.Navigate(webEmailPage);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
        }

        private void Refresh_OnClick()
        {
            try
            {
                _browser.Navigate(Address);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }

        }

        private void TbUrl_OnKeyDown(string  address)
        {
            _browser.Navigate(!string.IsNullOrEmpty(address) ? address : Address);
        }

        private void Back_OnClick()
        {
            _browser.GoBack();
        }

        private void Forward_OnClick()
        {
            _browser.GoForward();
        }

        private void New_OnClick()
        {
            var element = document.GetElementById("newmsgc");
            if (element != null)
            {
                element.InvokeMember("click");
            }
        }
        private void LogIn_OnClick()
        {
            LogInLayoutVisibility = Visibility.Collapsed;
            MainLayoutVisibility = Visibility.Visible;
        }

        private void LogOut_OnClick()
        {
            LogInLayoutVisibility = Visibility.Visible;
            MainLayoutVisibility = Visibility.Collapsed;
        }

        #endregion

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

            if (element != null)
                element.text = JavaScripts.AlertBlocker + " " + JavaScripts.OnbeforeunloadBlocker;

            head?.AppendChild(scriptEl);
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
       
    }
}
