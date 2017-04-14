using System;
using System.Windows;
using log4net;

namespace DYL.EmailIntegration
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public App()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += MyHandler;
        }

        static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception ex = (Exception)args.ExceptionObject;
            Log.Fatal("Unhandeled exception happened in application.", ex);
            MessageBox.Show("Unhandeled exception happened in application. See log file for details.");
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            log4net.Config.XmlConfigurator.Configure();
            Log.Info("Application Started");
            base.OnStartup(e);
        }  
    }
}
