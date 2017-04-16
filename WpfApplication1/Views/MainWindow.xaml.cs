using System;
using System.ComponentModel;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using DYL.EmailIntegration.ViewModels;
using log4net;
using mshtml;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace DYL.EmailIntegration.Views
{       
    /// <summary>
    /// Interaction logic for MainViewModel.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel(Browser);
        }
    }

}
