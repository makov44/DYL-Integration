using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Threading;
using DYL.EmailIntegration.Domain;
using DYL.EmailIntegration.Domain.Data;
using DYL.EmailIntegration.Helpers;
using DYL.EmailIntegration.Models;
using DYL.EmailIntegration.UI.Properties;
using log4net;
using mshtml;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using WebBrowser = System.Windows.Forms.WebBrowser;

namespace DYL.EmailIntegration.ViewModels
{
    public class MainViewModel: ViewModelBase
    {
        #region commands
        public ICommand LogInCommand => new DelegateCommand(LogIn_OnClick);
        public ICommand LogOutCommand => new DelegateCommand(param => { LogOut_OnClick(); });
        public ICommand BackCommand => new DelegateCommand(param => { Back_OnClick();});
        public ICommand ForwardCommand => new DelegateCommand(param => { Forward_OnClick();});
        public ICommand SendEmailsCommand => new DelegateCommand(param=> { SendEmails_OnClick(); });
        public ICommand SendCommand => new DelegateCommand(param => { Send_OnClick(); });
        public ICommand DeleteCommand => new DelegateCommand(param => { Delete_OnClick(); });
        public ICommand DeleteAllCommand => new DelegateCommand(param => { DeleteAll_OnClick(); });
        public ICommand RefreshCommand => new DelegateCommand(param => { Refresh_OnClick(); });
        public ICommand EnterCommand => new DelegateCommand(param => { TbUrl_OnKeyDown(param?.ToString()); });
        #endregion

        #region properties
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly WebBrowser _browser;
        private HtmlDocument _document;
        private readonly string _startPage =Settings.Default.AllStatesGatewayUrl;
        private readonly string _outlookHomePage = Settings.Default.AllStatesOutlookUrl;
        private PageName _currentPage;
        private readonly ConcurrentQueue<Email> _currentEmailQueue = new ConcurrentQueue<Email>();
        private int _remainingEmails;
        public int RemainingEmails
        {
            get { return _remainingEmails; }
            set
            {
                _remainingEmails = value;
                RaisePropertyChangedEvent("RemainingEmails");
            }
        }

        private int _sentEmails;
        public int SentEmails
        {
            get { return _sentEmails; }
            set
            {
                _sentEmails = value;
                RaisePropertyChangedEvent("SentEmails");
            }
        }

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

        private string _userName;
        public string UserName
        {
            get { return _userName; }
            set
            {
                _userName = value;
                RaisePropertyChangedEvent("UserName");
            }
        }

        private Visibility _emailsListLayout;
        public Visibility EmailsListLayout
        {
            get { return _emailsListLayout; }
            set
            {
                _emailsListLayout = value;
                RaisePropertyChangedEvent("EmailsListLayout");
            }
        }

        private Visibility _emailLayout;
        public Visibility EmailLayout
        {
            get { return _emailLayout; }
            set
            {
                _emailLayout = value;
                RaisePropertyChangedEvent("EmailLayout");
            }
        }

        private Visibility _isLoginInvalid;
        public Visibility IsLoginInvalid
        {
            get { return _isLoginInvalid; }
            set
            {
                _isLoginInvalid = value;
                RaisePropertyChangedEvent("IsLoginInvalid");
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

        #endregion

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
            _browser.Navigate(_startPage);
            Context.EmailQueue.CollectionChanged += EmailQueue_CollectionChanged;
            Context.PropertyChanged += Context_PropertyChanged;
            _currentPage = PageName.None;
            SentEmails = 0;
            LogInLayoutVisibility = Visibility.Visible;
            MainLayoutVisibility = Visibility.Collapsed;
            IsLoginInvalid = Visibility.Collapsed;
            ShowEmailPage(false);
        }

        private void Context_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "Session")
                return;

            if (Context.Session == null)
            {
                LogInLayoutVisibility = Visibility.Visible;
                MainLayoutVisibility = Visibility.Collapsed;
            }
            else
            {
                LogInLayoutVisibility = Visibility.Collapsed;
                MainLayoutVisibility = Visibility.Visible;
                ApplicationService.GetEmails(Context.Session.Key);
            }
        }

        private void EmailQueue_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            if(e.Action == NotifyCollectionChangedAction.Add)
                RemainingEmails ++;
            else if (e.Action == NotifyCollectionChangedAction.Remove)
                RemainingEmails--;
        }

        #region Browser events handlers
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
            _document = _browser.Document;
            if (_document != null && _document.Window != null)
                _document.Window.Error += Window_Error;
            DisableAlertWindow();
            if(_currentPage == PageName.New)
                NewEmailHandler();
            else if(_currentPage == PageName.Home)
                HomeEmailHandler();

            _currentPage = PageName.None;
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
                url = _document?.Url?.ToString();
            }

            e.Handled = true;
            Log.Error($"Document error happened. Error:{e.Description}, lineNumber:{e.LineNumber}, url:{url}");
        }

        private void Send_OnClick()
        {
            try
            {
                var sendButton = _document.GetElementById("send");
                if (sendButton == null)
                {
                    Log.Error("Can't find element by id 'send' on the page.");
                    return;
                }
                sendButton.Focus();
                sendButton.InvokeMember("click");
                _currentPage = PageName.Home;
                _browser.Navigate(_outlookHomePage);
                SentEmails++;
               PostEmailStatus("sent");
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                PostEmailStatus("failed");
            }
        }

        private void PostEmailStatus(string status)
        {
            Email email;
            _currentEmailQueue.TryDequeue(out email);

            if (email == null)
                return;

            ApplicationService.PostEmailStatus(new Status
            {
                EmailId = email.Id,
                StatusName = status
            });
        }

        private void Delete_OnClick()
        {
            try
            {
                PostEmailStatus("deleted");
                _currentPage = PageName.Home;
                _browser.Navigate(_outlookHomePage);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
        }

        private void DeleteAll_OnClick()
        {
            if (Context.EmailQueue.IsEmpty)
                return;

            var result = MessageBox.Show("Are you sure you want to delete all emails?", "Warning", MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (result != MessageBoxResult.OK)
                return;

            lock (Context.EmailQueue)
            {
                var count = Context.EmailQueue.Count;
                for (var i = 0; i < count; i++)
                {
                    Email email;
                    Context.EmailQueue.TryDequeue(out email);
                    if (email == null)
                        continue;

                    ApplicationService.PostEmailStatus(new Status
                    {
                        EmailId = email.Id,
                        StatusName = "deleted"
                    });
                }
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

        private void SendEmails_OnClick()
        {
            if (Context.EmailQueue.Count <= 0)
            {
                MessageBox.Show("There are no emails to send.", "Warning", MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var element = _document.GetElementById("newmsgc");
            if (element != null)
            {
                _currentPage = PageName.New;
                element.InvokeMember("click");
            }
            else
            {
                Log.Error("Can't find element 'newmsgc' on the page.");
            }
        }
        private void LogIn_OnClick(object parameter)
        {
            var passwordBox = parameter as PasswordBox;
            var password = passwordBox?.Password;
            var credentials = new Credentials
            {
                email = UserName,
                password = password
            };

            ApplicationService.GetNewSessionKey(credentials, key =>
            {
                if (string.IsNullOrEmpty(key))
                {
                    IsLoginInvalid = Visibility.Visible;
                    Log.Error("Failed to login.");
                }

                Context.Session = !string.IsNullOrEmpty(key) ? new Session(key, DateTime.Now) : null;

                if (!string.IsNullOrEmpty(key))
                    Authentication.SaveCredentials(credentials);
            });
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
            var script = _document?.GetElementById(scriptId);

            if (script != null)
                return;

            HtmlElement head = _document?.GetElementsByTagName("head")[0];
            HtmlElement scriptEl = _document?.CreateElement("script");
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
            ShowEmailPage(true);
            Email email;
            Context.EmailQueue.TryDequeue(out email);
            if (email == null)
            {
               Log.Error("Can't retrieve email from the queue.");
                return;
            }

            ClearCurrentEmailQueue();

            _currentEmailQueue.Enqueue(email);

            var frame = _document.GetElementById("ifBdy");
            var txtSubj = _document.GetElementById("txtSubj");
            var divTo = _document.GetElementById("divTo");

            if (txtSubj == null || divTo == null || frame == null)
            {
                Log.Debug($"It is not a New email page. txtSubj={txtSubj}," +
                          $" divTo={divTo}, frame={frame}, page url={_document.Url}");
                return;
            }

            if (txtSubj.GetAttribute("value") != email.Subject)
                txtSubj.SetAttribute("value", email.Subject);

            if (divTo.InnerText != email.To)
                divTo.InnerText = email.To;

            var scriptId = "replaceIFrameContentIId";
            var script = _document.GetElementById(scriptId);
            if (script != null)
            {
                _document.InvokeScript("replaceIFrameContent");
                return;
            }

            var head = _document?.GetElementsByTagName("head")[0];
            var scriptEl = _document?.CreateElement("script");
            var element = (IHTMLScriptElement)scriptEl?.DomElement;

            if (scriptEl == null || element == null)
                return;

            element.text = JavaScripts.ReplaceIFrameContent.Replace("=eewwfdfadsdffgnvbnbvhkiussdavcvbgfhyt=", email.Body);
            scriptEl.Id = scriptId;
            head?.AppendChild(scriptEl);

            _document.InvokeScript("replaceIFrameContent");
        }

        private void ClearCurrentEmailQueue()
        {
            if (_currentEmailQueue.IsEmpty)
                return;

            lock (_currentEmailQueue)
            {
                var count = _currentEmailQueue.Count;
                for (var i = 0; i < count; i++)
                {
                    Email oldEmail;
                    _currentEmailQueue.TryDequeue(out oldEmail);
                    Log.Error($"Current email queue is not empty, possible race condition. EmailId={oldEmail.Id}");
                }
            }
        }

        private void HomeEmailHandler()
        {
            
            if (Context.EmailQueue.Count > 0)
            {
                Task.Run(()=>
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        Task.Delay(1000);
                        SendEmails_OnClick();
                    });
                });
            }
            else
            {
                ShowEmailPage(false);
            }
        }

        private void ShowEmailPage(bool flag)
        {
            if (flag)
            {
                EmailLayout = Visibility.Visible;
                EmailsListLayout = Visibility.Collapsed;
            }
            else
            {
                EmailLayout = Visibility.Collapsed;
                EmailsListLayout = Visibility.Visible;
            }
        }
    }
}
