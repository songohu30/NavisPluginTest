using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestNavisPlugin
{
    public class ViewModel : INotifyPropertyChanged
    {
        private string _logInfo;
        public string LogInfo
        {
            get => _logInfo;
            set
            {
                if (_logInfo != value)
                {
                    _logInfo = value;
                    OnPropertyChanged(nameof(LogInfo));
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ViewModel()
        {
            LogInfo = "Initialized new dimension tool..." + Environment.NewLine;
        }
    }
}
