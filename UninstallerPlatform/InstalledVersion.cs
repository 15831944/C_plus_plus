#define DEBUG

using System;
using System.ComponentModel;
using System.Globalization;
using System.Management;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;


namespace DoxFlow.Platform1C
{
    internal class InstalledVersion : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string name;
        private string version;
        private bool isChecked;
        private State state;
        private string uninstallStr;

        public InstalledVersion(string Name, string Version, bool IsChecked, State State, string uninstallString)
        {
            this.name = Name;
            this.version = Version;
            this.isChecked = IsChecked;
            this.state = State;
            this.uninstallStr = uninstallString;
        }

        #region Property

        public string Name
        {
            get { return name; }
            set
            {
                if (name != value)
                {
                    name = value;
                    OnPropertyChanged("Name");
                }
            }
        }
        public string Version
        {
            get { return version; }
            set
            {
                if (version != value)
                {
                    version = value;
                    OnPropertyChanged("Version");
                }
            }
        }
        public bool IsChecked
        {
            get { return isChecked; }
            set
            {
                if (isChecked != value)
                {
                    isChecked = value;
                    OnPropertyChanged("IsChecked");
                }
            }
        }
        public State State
        {
            get { return state; }
            set
            {
                if (state != value)
                {
                    state = value;
                    OnPropertyChanged("State");
                }
            }
        }

        #endregion

        protected virtual void OnPropertyChanged(string propChanged)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propChanged));
        }

        public Task<bool> Uninstall()
        {
            return Task.Run(() =>
            {
                try
                {
                    if (uninstallStr == null)
                    {
                        return false;
                    }

                    var ProcessInfo = new System.Diagnostics.ProcessStartInfo("msiexec.exe", String.Format("/X{0} /quiet", uninstallStr));
                    ProcessInfo.CreateNoWindow = true;
                    ProcessInfo.UseShellExecute = true;

                    var Process = System.Diagnostics.Process.Start(ProcessInfo);
                    Process.WaitForExit();
                    var result = Process.ExitCode;
                    Process.Close();

                    if (result == 0)
                    {
                        State = State.UnInstalled;
                        IsChecked = false;

                        return true;
                    }
                    else
                        return false;

                }
                catch (Exception)
                {
                    return false;
                }
            });
        }
    }

    enum State
    {
        Installed,
        UnInstalled,
        Uninstalling
    }

    [ValueConversion(typeof(State), typeof(FontWeights))]
    public class StateToFontWeightsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if ((State)value == State.Uninstalling)
                return FontWeights.Bold;
            else
                return FontWeights.Normal;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException("This method should never be called");
        }
    }
}
