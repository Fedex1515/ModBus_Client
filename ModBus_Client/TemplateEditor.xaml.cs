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

using System.Collections.ObjectModel;

// Json lib
using System.Web.Script.Serialization;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for TemplateEditor.xaml
    /// </summary>
    public partial class TemplateEditor : Window
    {
        MainWindow main;
        String pathToConfiguration = "";

        ObservableCollection<ModBus_Item> list_coilsTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_inputsTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_inputRegistersTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_holdingRegistersTable = new ObservableCollection<ModBus_Item>();

        public TemplateEditor(MainWindow main_)
        {
            InitializeComponent();

            main = main_;
            pathToConfiguration = main.pathToConfiguration;

            dataGridViewCoils.ItemsSource = list_coilsTable;
            dataGridViewInput.ItemsSource = list_inputsTable;
            dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;
            dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

            main.Dispatcher.Invoke((Action)delegate
            {
                if (main.tabControlMain.SelectedIndex > 0 && main.tabControlMain.SelectedIndex < 5)
                {
                    tabControlTemplate.SelectedIndex = main.tabControlMain.SelectedIndex - 1;
                }
            });
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            salvaTemplate();

            if (MessageBox.Show("Ricaricare il template nel client?", "Info", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
            {
                main.salva_configurazione(false);
                main.carica_configurazione();
            }
        }

        private bool salvaTemplate()
        {
            try
            {
                var template = new TEMPLATE();

                template.comboBoxCoilsOffset_ = comboBoxCoilsOffset.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputOffset_ = comboBoxInputOffset.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputRegOffset_ = comboBoxInputRegOffset.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxHoldingOffset_ = comboBoxHoldingOffset.SelectedIndex == 1 ? "HEX" : "DEC";

                template.comboBoxCoilsRegistri_ = comboBoxCoilsRegistri.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputRegistri_ = comboBoxInputRegistri.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputRegRegistri_ = comboBoxInputRegRegistri.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxHoldingRegistri_ = comboBoxHoldingRegistri.SelectedIndex == 1 ? "HEX" : "DEC";

                template.textBoxCoilsOffset_ = textBoxCoilsOffset.Text;
                template.textBoxInputOffset_ = textBoxInputOffset.Text;
                template.textBoxInputRegOffset_ = textBoxInputRegOffset.Text;
                template.textBoxHoldingOffset_ = textBoxHoldingOffset.Text;

                template.dataGridViewCoils = list_coilsTable.ToArray<ModBus_Item>();
                template.dataGridViewInput = list_inputsTable.ToArray<ModBus_Item>();
                template.dataGridViewInputRegister = list_inputRegistersTable.ToArray<ModBus_Item>();
                template.dataGridViewHolding = list_holdingRegistersTable.ToArray<ModBus_Item>();

                JavaScriptSerializer jss = new JavaScriptSerializer();
                string file_content = jss.Serialize(template);

                File.WriteAllText("Json/" + pathToConfiguration + "/Template.json", file_content);

                Console.WriteLine("Caricata configurazione precedente\n");

                return true;
            }
            catch (Exception err)
            {
                Console.WriteLine("Errore caricamento configurazione\n");
                Console.WriteLine(err);

                return false;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/Template.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                TEMPLATE template = jss.Deserialize<TEMPLATE>(file_content);

                comboBoxCoilsOffset.SelectedIndex = template.comboBoxCoilsOffset_ == "HEX" ? 1 : 0;
                comboBoxInputOffset.SelectedIndex = template.comboBoxInputOffset_ == "HEX" ? 1 : 0;
                comboBoxInputRegOffset.SelectedIndex = template.comboBoxInputRegOffset_ == "HEX" ? 1 : 0;
                comboBoxHoldingOffset.SelectedIndex = template.comboBoxHoldingOffset_ == "HEX" ? 1 : 0;

                comboBoxCoilsRegistri.SelectedIndex = template.comboBoxCoilsRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegistri.SelectedIndex = template.comboBoxInputRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegRegistri.SelectedIndex = template.comboBoxInputRegRegistri_ == "HEX" ? 1 : 0;
                comboBoxHoldingRegistri.SelectedIndex = template.comboBoxHoldingRegistri_ == "HEX" ? 1 : 0;

                textBoxCoilsOffset.Text = template.textBoxCoilsOffset_;
                textBoxInputOffset.Text = template.textBoxInputOffset_;
                textBoxInputRegOffset.Text = template.textBoxInputRegOffset_;
                textBoxHoldingOffset.Text = template.textBoxHoldingOffset_;

                // Tabella coils
                foreach (ModBus_Item item in template.dataGridViewCoils)
                {
                    list_coilsTable.Add(item);
                }

                // Tabella inputs
                foreach (ModBus_Item item in template.dataGridViewInput)
                {
                    list_inputsTable.Add(item);
                }

                // Tabella input registers
                foreach (ModBus_Item item in template.dataGridViewInputRegister)
                {
                    list_inputRegistersTable.Add(item);
                }

                // Tabella holdings
                foreach (ModBus_Item item in template.dataGridViewHolding)
                {
                    list_holdingRegistersTable.Add(item);
                }

                Console.WriteLine("Caricata configurazione precedente\n");
            }
            catch(Exception err)
            {
                Console.WriteLine("Errore caricamento configurazione\n");
                Console.WriteLine(err);
            }
        }

        private void salvaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (salvaTemplate())
            {
                MessageBox.Show("Template salvato correttamente", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Errore salvataggio template", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
            {
                this.Close();
            }
        }
    }

    // Classe per caricare dati dal file di configurazione json
    public class TEMPLATE
    {
        public string comboBoxCoilsOffset_ { get; set; }
        public string comboBoxInputOffset_ { get; set; }
        public string comboBoxInputRegOffset_ { get; set; }
        public string comboBoxHoldingOffset_ { get; set; }

        public string comboBoxCoilsRegistri_ { get; set; }
        public string comboBoxInputRegistri_ { get; set; }
        public string comboBoxInputRegRegistri_ { get; set; }
        public string comboBoxHoldingRegistri_ { get; set; }

        public string textBoxCoilsOffset_ { get; set; }
        public string textBoxInputOffset_ { get; set; }
        public string textBoxInputRegOffset_ { get; set; }
        public string textBoxHoldingOffset_ { get; set; }

        public ModBus_Item[] dataGridViewCoils { get; set; }
        public ModBus_Item[] dataGridViewInput { get; set; }
        public ModBus_Item[] dataGridViewInputRegister { get; set; }
        public ModBus_Item[] dataGridViewHolding { get; set; }
    }

    public class ModBus_Item
    {
        public string Register { get; set; }
        public string Value { get; set; }
        public string ValueBin { get; set; }
        public string Notes { get; set; }
        public string Color { get; set; }
    }
}
