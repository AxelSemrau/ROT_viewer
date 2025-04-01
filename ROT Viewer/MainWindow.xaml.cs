using System;
using System.Threading;
using System.Windows;
using System.Windows.Data;

namespace ROT_Viewer
{
    /// <summary>
    /// Shows all objects in the COM running object table.
    /// </summary>
    public partial class MainWindow : Window
    {
        // ReSharper disable once NotAccessedField.Local
        private readonly Timer mUpdateTimer;

        public MainWindow()
        {
            InitializeComponent();
            mUpdateTimer = new Timer(OnTimer, null, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1));
        }

        private void OnTimer(object state)
        {
            Dispatcher.Invoke(() => CollectionViewSource.GetDefaultView(LvEntries.ItemsSource).Refresh());
        }
    }
}
