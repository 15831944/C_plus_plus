using DoxFlow.Platform1C;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Shell;

namespace DoxFlow.Uninstaller1C
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<InstalledVersion> InstalVer;
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void btnUninstall_Click(object sender, RoutedEventArgs e)
        {
            list.IsEnabled = false;
            btnUninstall.IsEnabled = false;
            btnUninstall.Content = "Uninstalling...";

            List<InstalledVersion> uninst = new List<InstalledVersion>();
            foreach (var item in InstalVer)
            {
                if (item.IsChecked && item.State == State.Installed)
                {
                    uninst.Add(item);
                }
            }

            int ind = 0;
            this.TaskbarItemInfo.ProgressValue = 0d;
            this.TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;

            foreach (var item in uninst)
            {
                if (item.IsChecked && item.State == State.Installed)
                {
                    item.State = State.Uninstalling;
                    this.TaskbarItemInfo.Description = item.Name + " (версия: " + item.Version + ")";

                    var result = await item.Uninstall();
                    if (result)
                    {
                        InstalVer.Remove(item);
                        ind++;
                        this.TaskbarItemInfo.ProgressValue = (double)ind / (double)uninst.Count;
                    }
                }
            }

            list.IsEnabled = true;
            btnUninstall.IsEnabled = true;
            btnUninstall.Content = "Uninstall";
            this.TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            InstalVer = new ObservableCollection<InstalledVersion>(InstalledVersionCollection.GetVersions());

            Binding binding = new Binding();
            binding.Source = InstalVer;
            list.SetBinding(ListView.ItemsSourceProperty, binding);
        }

        private void list_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            foreach (InstalledVersion item in ((System.Windows.Controls.ListBox)(sender)).SelectedItems)
            {
                if (e.Key == System.Windows.Input.Key.Space)
                {
                    item.IsChecked = !item.IsChecked;
                }
            }
        }
    }
}
