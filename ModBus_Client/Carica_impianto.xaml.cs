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

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for Carica_impianto.xaml
    /// </summary>
    public partial class Carica_impianto : Window
    {
        public string path { get; set; }

        dynamic languageTemplate;

        public Carica_impianto(String defaultPath, MainWindow main_)
        {
            InitializeComponent();

            String[] subFolders = Directory.GetDirectories("Json\\");

            comboBoxImpianto.Items.Clear();

            for (int i = 0; i < subFolders.Length; i++)
                comboBoxImpianto.Items.Add(subFolders[i].Substring(subFolders[i].IndexOf('\\') + 1));

            comboBoxImpianto.SelectedItem = defaultPath;

            loadLanguageTemplate(main_.language);
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            path = comboBoxImpianto.SelectedItem.ToString();
            this.DialogResult = true;
        }

        private void buttonAnnulla_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
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

        private void LodaProfile_Loaded(object sender, RoutedEventArgs e)
        {
            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);

            comboBoxImpianto.Focus();
        }

        private void LodaProfile_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                buttonOpen_Click(sender, e);
            }

            if(e.Key == Key.Escape)
            {
                buttonAnnulla_Click(sender, e);
            }
        }
    }
}
