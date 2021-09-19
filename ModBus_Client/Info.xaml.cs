using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Reflection;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for Info.xaml
    /// </summary>
    public partial class Info : Window
    {
        public Info(string title, string version)
        {
            InitializeComponent();

            var version_ = Assembly.GetEntryAssembly().GetName().Version;
            DateTime buildDateTime_ = new DateTime(2000, 1, 1).Add(new TimeSpan(
            TimeSpan.TicksPerDay * version_.Build + // days since 1 January 2000
            TimeSpan.TicksPerSecond * 2 * version_.Revision)); // seconds since midnight, (multiply by 2 to get original)

            labelAuthor.Content = "Federico Turco";
            labelVersion.Content = title + " - " + version;

            labelBuildNumber.Content = version_;
            LabelBuildDate.Content = buildDateTime_.ToString("dd / MM / yyyy - HH : mm : ss");

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }
        }
    }
}
