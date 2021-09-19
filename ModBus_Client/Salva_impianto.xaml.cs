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
using System.IO;

// Json lib
using System.Web.Script.Serialization;

// Libreria lingue
using LanguageLib; // Libreria custom per caricare etichette in lingue differenti

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for Salva_impianto.xaml
    /// </summary>
    public partial class Salva_impianto : Window
    {
        public string path { get; set; }

        Language lang;

        public Salva_impianto(MainWindow main_)
        {
            InitializeComponent();

            lang = new Language(this);
            lang.loadLanguageTemplate(main_.language);

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
        }

        private void buttonAnnulla_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            path = textBoxImpianto.Text.ToString();
            this.DialogResult = true;
        }

        private void SaveProfile_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                buttonOpen_Click(sender, e);
            }

            if (e.Key == Key.Escape)
            {
                buttonAnnulla_Click(sender, e);
            }
        }

        private void SaveProfile_Loaded(object sender, RoutedEventArgs e)
        {
            textBoxImpianto.Focus();
        }

        private void textBoxImpianto_KeyUp(object sender, KeyEventArgs e)
        {
            // Commentato perchè crasha premdendo invio per salvare
            /*if(e.Key == Key.Enter)
            {
                buttonOpen_Click(sender, e);
            }*/
        }
    }
}
