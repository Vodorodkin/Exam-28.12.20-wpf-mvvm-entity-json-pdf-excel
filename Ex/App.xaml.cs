using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Ex.Views;
using Ex.ViewsModels;

namespace Ex
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            new MainWindow()
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Title = "Главная форма",
                DataContext = new VM_MainWindow()
                {

                }
            }.Show();
        }
    }
}
