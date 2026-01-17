using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace NetIsolatePlus
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
            var v = Assembly.GetExecutingAssembly().GetName().Version;
            VersionText.Text = v is null ? "" : $"Version {v.Major}.{v.Minor}.{v.Build}";
            PreviewKeyDown += (_, e) => { if (e.Key == Key.Escape) Close(); };
        }

        private void DragWindow(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
