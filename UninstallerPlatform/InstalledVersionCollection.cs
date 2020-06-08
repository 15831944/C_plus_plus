using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Windows.Threading;

namespace DoxFlow.Platform1C
{
    internal class InstalledVersionCollection
    {
        internal static List<InstalledVersion> GetVersions()
        {
            var list = new List<InstalledVersion>();

            var subkey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            foreach (var itemUUID in subkey.GetSubKeyNames())
            {
                var itemKey = subkey.OpenSubKey(itemUUID);
                string name = (string)itemKey.GetValue("DisplayName", null);
                if (name != null)
                {
                    string vendor = (string)itemKey.GetValue("Publisher", null);
                    if (vendor == "1C" || vendor == "1С")
                    {
                        string version = (string)itemKey.GetValue("DisplayVersion", "0.0.0.0");
                        InstalledVersion instVerItem = new InstalledVersion(name, version, false, State.Installed, itemUUID);
                        list.Add(instVerItem);
                    }
                }
            }

            return list.OrderBy(x => x.Version).ToList<InstalledVersion>();
        }
    }
}
