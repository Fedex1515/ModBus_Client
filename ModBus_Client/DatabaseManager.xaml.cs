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
using System.Collections.ObjectModel;
using System.IO;
using System.Diagnostics;
using System.IO.Compression;
using Microsoft.Win32;

// Libreria lingue
using LanguageLib; // Libreria custom per caricare etichette in lingue differenti

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for DatabaseManager.xaml
    /// </summary>
    public partial class DatabaseManager : Window
    {
        ObservableCollection<profile> db = new ObservableCollection<profile>();

        Language lang;

        public DatabaseManager(MainWindow main_)
        {
            InitializeComponent();

            lang = new Language(this);

            DataGridDb.ItemsSource = db;

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);

            lang.loadLanguageTemplate(main_.language);

            this.Title = main_.title + " " + main_.version;
        }

         void LoadDb()
        {
            String[] subFolders = Directory.GetDirectories("Json\\");

            db.Clear();

            for (int i = 0; i < subFolders.Length; i++)
            {
                profile tmp = new profile();

                tmp.name = subFolders[i].Substring(subFolders[i].IndexOf('\\') + 1);

                db.Add(tmp);
            }
        }

        private void ButtonRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadDb();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDb();
        }

        private void ButtonOpenFileLocation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("explorer.exe", "Json");
            }
            catch
            {
                MessageBox.Show("Imposssibile aprire la cartella di configurazione Database", "Alert");
                Console.WriteLine("Imposssibile aprire la cartella di configurazione Database");
            }
        }

        private void ButtonExportZip_Click(object sender, RoutedEventArgs e)
        {
            profile currItem = (profile)DataGridDb.SelectedItem;
            
            SaveFileDialog window = new SaveFileDialog();

            window.Filter = "Zip Files | *.zip";
            window.DefaultExt = ".zip";
            window.FileName = currItem.name + ".zip";

            if ((bool)window.ShowDialog())
            {
                ZipFile.CreateFromDirectory("Json\\" + currItem.name, window.FileName);
            }
        }

        private void ButtonImportZip_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog window = new OpenFileDialog();

            window.Filter = "Zip Files | *.zip";
            window.DefaultExt = ".zip";

            if ((bool)window.ShowDialog())
            {
                ZipFile.ExtractToDirectory(window.FileName, "Json\\" + System.IO.Path.GetFileNameWithoutExtension(window.FileName));
            }

            LoadDb();
        }

        private void DataGridDb_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                profile selected = (profile)DataGridDb.SelectedItem;

                labelProfileSelected.Content = labelProfileSelected.Content.ToString().Split(':')[0] + ": " + selected.name;
                labelProfileSelected.Visibility = Visibility.Visible;

                ButtonExportZip.IsEnabled = true;
                //ButtonImportZip.IsEnabled = true;
            }
            catch
            {
                ButtonExportZip.IsEnabled = false;
                //ButtonImportZip.IsEnabled = false;
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
            {
                this.Close();
            }
        }
    }

    public class profile
    {
        public String name { get; set; }
    }
}
