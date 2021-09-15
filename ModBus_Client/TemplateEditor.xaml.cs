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
using Microsoft.Win32;

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

        dynamic languageTemplate;

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

            loadLanguageTemplate(main.language);
        }

        public void loadLanguageTemplate(string templateName)
        {
            string file_content = File.ReadAllText("Lang/" + templateName + ".json");

            JavaScriptSerializer jss = new JavaScriptSerializer();
            languageTemplate = jss.Deserialize<dynamic>(file_content);

            foreach (KeyValuePair<string, dynamic> group in languageTemplate)
            {
                // debug
                // Console.WriteLine("Group: " + group.Key + ": " + group.Value);

                switch (group.Key)
                {
                    // LABELS
                    case "labels":
                        // debug
                        // Console.WriteLine("Fould label group");

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            // debug
                            // Console.WriteLine("Value: " + item.Key + ": " + item.Value);

                            var myLabel = (Label)this.FindName(item.Key);

                            if (myLabel != null)
                            {
                                myLabel.Content = item.Value;
                            }
                        }
                        break;

                    // RADIOBUTTONS
                    case "radioButtons":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myRadioButton = (RadioButton)this.FindName(item.Key);

                            if (myRadioButton != null)
                            {
                                myRadioButton.Content = item.Value;
                            }
                        }
                        break;

                    // CHECKBOXES
                    case "checkBoxes":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myCheckBox = (CheckBox)this.FindName(item.Key);

                            if (myCheckBox != null)
                            {
                                myCheckBox.Content = item.Value;
                            }
                        }
                        break;

                    // BUTTONS
                    case "buttons":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myButton = (Button)this.FindName(item.Key);

                            if (myButton != null)
                            {
                                myButton.Content = item.Value;
                            }
                        }
                        break;

                    // TABITEMS
                    case "tabItems":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myTab = (TabItem)this.FindName(item.Key);

                            if (myTab != null)
                            {
                                myTab.Header = item.Value;
                            }
                        }
                        break;

                    // MENUITEMS
                    case "menuItems":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            var myMenu = (MenuItem)this.FindName(item.Key);

                            if (myMenu != null)
                            {
                                myMenu.Header = item.Value;
                            }
                        }
                        break;

                    // TOOLTIPS
                    case "toolTips":

                        foreach (KeyValuePair<string, dynamic> item in group.Value)
                        {
                            // debug
                            object obj = this.FindName(item.Key);

                            /*try
                            {
                                // Button
                                if (obj.ToString().Split(' ')[0].IndexOf("System.Windows.Controls.Button") != -1)
                                {
                                    var myButton = (Button)this.FindName(item.Key);

                                    if (myButton != null)
                                    {
                                        myButton.ToolTip = item.Value;
                                    }
                                }

                                // Label
                                if (this.FindName(item.Key).ToString().Split(' ')[0].IndexOf("System.Windows.Controls.Label") != -1)
                                {
                                    var myLabel = (Label)this.FindName(item.Key);

                                    if (myLabel != null)
                                    {
                                        myLabel.ToolTip = item.Value;
                                    }
                                }

                                // CheckBox
                                if (this.FindName(item.Key).ToString().Split(' ')[0].IndexOf("System.Windows.Controls.CheckBox") != -1)
                                {
                                    var myCheck = (CheckBox)this.FindName(item.Key);

                                    if (myCheck != null)
                                    {
                                        myCheck.ToolTip = item.Value;
                                    }
                                }
                            }
                            catch(Exception err)
                            {
                                Console.WriteLine(err);
                            }*/
                        }
                        break;

                    // STRING
                    case "strings":
                        ;
                        break;
                }



            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            salvaTemplate();

            main.salva_configurazione(false);
            main.carica_configurazione();
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
            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);

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

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if(Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                switch (e.Key)
                {
                    case Key.D1:
                        tabControlTemplate.SelectedIndex = 0;
                        break;

                    case Key.D2:
                        tabControlTemplate.SelectedIndex = 1;
                        break;

                    case Key.D3:
                        tabControlTemplate.SelectedIndex = 2;
                        break;

                    case Key.D4:
                        tabControlTemplate.SelectedIndex = 3;
                        break;

                    case Key.D5:
                        tabControlTemplate.SelectedIndex = 4;
                        break;
                    
                    case Key.S:
                        salvaTemplate();
                        break;
                }
            }
        }

        private void ButtonImportCsvCoils_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_coilsTable);
        }

        private void ButtonImportCsvInputs_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_inputsTable);
        }

        private void ButtonImportCsvInputRegisters_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_inputRegistersTable);
        }

        private void ButtonImportCsvHoldingRegisters_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_inputRegistersTable);
        }

        private void ButtonImportCsvHoldingRegisters_Click_1(object sender, RoutedEventArgs e)
        {
            importCsv(list_holdingRegistersTable);
        }

        private void ButtonExportCsvCoils_Click(object sender, RoutedEventArgs e)
        {
            exportCsv(list_coilsTable, "_Coils");
        }

        private void ButtonExportCsvInputs_Click(object sender, RoutedEventArgs e)
        {
            exportCsv(list_inputsTable, "_Inputs");
        }

        private void ButtonExportCsvInputRegisters_Click(object sender, RoutedEventArgs e)
        {
            exportCsv(list_inputRegistersTable, "_InputRegisters");
        }

        private void ButtonExportCsvHoldingRegisters_Click(object sender, RoutedEventArgs e)
        {
            exportCsv(list_holdingRegistersTable, "_HoldingRegisters");
        }

        public void exportCsv(ObservableCollection<ModBus_Item> collection, String append)
        {
            SaveFileDialog window = new SaveFileDialog();

            window.Filter = "csv Files | *.csv";
            window.DefaultExt = ".csv";
            window.FileName = main.pathToConfiguration + append + ".csv";

            if ((bool)window.ShowDialog())
            {
                String content = "Register,Value,Notes,Mappings\n";

                foreach(ModBus_Item item in collection)
                {
                    if(item != null)
                    {
                        content += item.Register + "," + item.Value + "," + item.Notes + "," + item.Mappings + "\n";
                    }
                }

                File.WriteAllText(window.FileName, content);
            }
        }

        public void importCsv(ObservableCollection<ModBus_Item> collection)
        {
            OpenFileDialog window = new OpenFileDialog();

            window.Filter = "csv Files | *.csv";
            window.DefaultExt = ".csv";

            if ((bool)window.ShowDialog())
            {
                string content = File.ReadAllText(window.FileName);
                string[] splitted = content.Split('\n');

                for (int i = 1; i < splitted.Count(); i++)
                {
                    ModBus_Item item = new ModBus_Item();

                    try
                    {
                        item.Register = splitted[i].Split(',')[0];
                        //item.Register = splitted[i].Split(',')[1];
                        item.Notes = splitted[i].Split(',')[2];
                        item.Mappings = splitted[i].Split(',')[3];

                        collection.Add(item);
                    }
                    catch
                    {
                        //Console.WriteLine(err);
                    }
                }
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
        public string Mappings { get; set; }
        public string Color { get; set; }
    }
}
