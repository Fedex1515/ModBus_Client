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

using System.Threading;

using ModBusMaster_Chicco;

using System.IO;

using Raccolta_funzioni_parser;

//Libreria JSON
using System.Web.Script.Serialization;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for Editor.xaml
    /// </summary>
    public partial class Editor : Window
    {
        static int numberOfRegisters = 16;  // Numero di registri comandabili dal form

        ModBus_Chicco ModBus;   // Oggetto modbus passato dal MAIN form

        ComandiBit sim_Auge;   // Ricevo come ingresso il form principale per aggiornare i bit della grafica tramite una funzione pubblica

        Parser P = new Parser();

        int row = 0;      // Riga editata dal form "Sim AUGE"

        Border[] pictureBixRegisterBits = new Border[16];
        TextBox[] textBoxBitsLabel = new TextBox[16];

        // Variabili per tenere in memoria le variabili del file "Configurazione_form_2.json"
        TextBox[] textBoxLabel = new TextBox[numberOfRegisters];
        TextBox[] textBoxRegister = new TextBox[numberOfRegisters];
        TextBox[] textBoxValue = new TextBox[numberOfRegisters];
        ComboBox[] comboBox = new ComboBox[numberOfRegisters];
        String[] labelBitRegisters = new String[16];

        String pathToConfiguration;

        public Editor(bool useOffsetInTextBox_, ModBus_Chicco ModBus_, int row_, ComandiBit sim_Auge_, String pathToConfiguration_, bool registriBloccati)
        {
            InitializeComponent();

            row = row_;

            // Creo evento di chiusura del form
            this.Closing += Editor_cs_Cosing;


            ModBus = ModBus_;
            sim_Auge = sim_Auge_;

            textBoxModBusAddress.Text = sim_Auge.textBoxModBusAddress.Text;

            pathToConfiguration = pathToConfiguration_;

            // Assegnazione array elementi grafici

            // Border icone bit del registro
            pictureBixRegisterBits[0] = pictureBox_0;
            pictureBixRegisterBits[1] = pictureBox_1;
            pictureBixRegisterBits[2] = pictureBox_2;
            pictureBixRegisterBits[3] = pictureBox_3;
            pictureBixRegisterBits[4] = pictureBox_4;
            pictureBixRegisterBits[5] = pictureBox_5;
            pictureBixRegisterBits[6] = pictureBox_6;
            pictureBixRegisterBits[7] = pictureBox_7;
            pictureBixRegisterBits[8] = pictureBox_8;
            pictureBixRegisterBits[9] = pictureBox_9;
            pictureBixRegisterBits[10] = pictureBox_10;
            pictureBixRegisterBits[11] = pictureBox_11;
            pictureBixRegisterBits[12] = pictureBox_12;
            pictureBixRegisterBits[13] = pictureBox_13;
            pictureBixRegisterBits[14] = pictureBox_14;
            pictureBixRegisterBits[15] = pictureBox_15;

            // Etichette dei bit
            textBoxBitsLabel[0] = textBoxLabel_0;
            textBoxBitsLabel[1] = textBoxLabel_1;
            textBoxBitsLabel[2] = textBoxLabel_2;
            textBoxBitsLabel[3] = textBoxLabel_3;
            textBoxBitsLabel[4] = textBoxLabel_4;
            textBoxBitsLabel[5] = textBoxLabel_5;
            textBoxBitsLabel[6] = textBoxLabel_6;
            textBoxBitsLabel[7] = textBoxLabel_7;
            textBoxBitsLabel[8] = textBoxLabel_8;
            textBoxBitsLabel[9] = textBoxLabel_9;
            textBoxBitsLabel[10] = textBoxLabel_10;
            textBoxBitsLabel[11] = textBoxLabel_11;
            textBoxBitsLabel[12] = textBoxLabel_12;
            textBoxBitsLabel[13] = textBoxLabel_13;
            textBoxBitsLabel[14] = textBoxLabel_14;
            textBoxBitsLabel[15] = textBoxLabel_15;

            textBoxHoldingOffset.Text = sim_Auge.textBoxHoldingOffset.Text;
            comboBoxHoldingOffset.SelectedIndex = sim_Auge.comboBoxHoldingOffset.SelectedIndex;

            textBoxVal_A.Text = sim_Auge.textBoxVal_A.Text;
            comboBoxVal_A.SelectedIndex = sim_Auge.comboBoxVal_A.SelectedIndex;

            textBoxLabel_A.IsEnabled = false;

            textBoxModBusAddress.IsEnabled = false;
            comboBoxHoldingOffset.IsEnabled = false;
            textBoxHoldingOffset.IsEnabled = false;
            comboBoxVal_A.IsEnabled = false;
            textBoxVal_A.IsEnabled = false;

            carica_configurazione();
            carica_configurazione_3();

            if (registriBloccati)
            {
                for (int i = 0; i < 16; i++)
                {
                    textBoxBitsLabel[i].IsEnabled = false;
                }
            }
        }

        public void Editor_cs_Cosing(object sender, EventArgs e)
        {
            salva_configurazione_3();
        }

        private void carica_configurazione()
        {
            //Caricamento valori ultima sessione da file del form "Sim_AUGE"
            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/ComandiBit.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                SAVE_Form config = jss.Deserialize<SAVE_Form>(file_content);

                textBoxModBusAddress.Text = config.textBoxModBusAddress_;
                textBoxHoldingOffset.Text = config.textBoxHoldingOffset_;
                comboBoxHoldingOffset.SelectedItem = config.comboBoxHoldingOffset_;

                // Variabili array 
                String[] textBoxLabel_ = new String[numberOfRegisters];
                String[] textBoxVal_ = new String[numberOfRegisters];
                String[] textBoxValue_ = new String[numberOfRegisters];

                String[,] labelBitRegisters_ = new String[numberOfRegisters, 16];

                object[] comboBoxVal_ = new object[numberOfRegisters];

                textBoxLabel_ = config.textBoxLabel_;
                textBoxVal_ = config.textBoxVal_;
                textBoxValue_ = config.textBoxValue_;
                comboBoxVal_ = config.comboBoxVal_;
                //labelBitRegisters_ = config.labelBitRegisters_;

                textBoxLabel_A.Text = textBoxLabel_[row];
                textBoxVal_A.Text = textBoxVal_[row];
                textBoxValue_A.Text = textBoxValue_[row];

                comboBoxVal_A.SelectedItem = comboBoxVal_[row];

                /*for (int i = 0; i < numberOfRegisters; i++)
                {
                    textBoxLabel[i].Text = textBoxLabel_[i];
                    textBoxRegister[i].Text = textBoxVal_[i];
                    textBoxValue[i].Text = textBoxValue_[i];

                    comboBox[i].SelectedItem = comboBoxVal_[i];
                }*/
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void carica_configurazione_3()
        {
            //Caricamento valori ultima sessione di questo form
            try
            {
                this.Title = sim_Auge.textBoxLabel[row].Text;
                textBoxLabel_A.Text = sim_Auge.textBoxLabel[row].Text;

                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/Label_ComandiBit_" + row.ToString() + ".json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                SAVE_Form3 config = jss.Deserialize<SAVE_Form3>(file_content);

                String[] labelBitRegisters_ = new String[16];

                labelBitRegisters_ = config.labelBitRegisters_;

                intToPicture(UInt16.Parse(sim_Auge.textBoxValue[row].Text));
                textBoxValue_A.Text = sim_Auge.textBoxValue[row].Text;
                textBoxVal_A.Text = sim_Auge.textBoxRegister[row].Text;

                for (int a = 0; a < 16; a++)
                {
                    labelBitRegisters[a] = labelBitRegisters_[a];

                    textBoxBitsLabel[a].Text = labelBitRegisters[a];
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        public void salva_configurazione_3()
        {
            //Caricamento valori ultima sessione di questo form
            try
            {
                SAVE_Form3 config = new SAVE_Form3();

                String[] labelBitRegisters_ = new String[16];

                for (int a = 0; a < 16; a++)
                {
                    labelBitRegisters_[a] = textBoxBitsLabel[a].Text;
                }

                config.labelBitRegisters_ = labelBitRegisters_;

                JavaScriptSerializer jss = new JavaScriptSerializer();
                string file_content = jss.Serialize(config);

                File.WriteAllText("Json/" + pathToConfiguration + "/Label_ComandiBit_" + row.ToString() + ".json", file_content);
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        /*public void salva_configurazione()
        {
            //Salvataggio valori ultima sessione
            try
            {
                SAVE_Form config = new SAVE_Form();

                config.textBoxModBusAddress_ = textBoxModBusAddress.Text;
                config.textBoxHoldingOffset_ = textBoxHoldingOffset.Text;
                config.comboBoxHoldingOffset_ = comboBoxHoldingOffset.SelectedItem;

                // Variabili array 
                String[] textBoxLabel_ = new String[numberOfRegisters];
                String[] textBoxVal_ = new String[numberOfRegisters];
                String[] textBoxValue_ = new String[numberOfRegisters];

                String[,] labelBitRegisters_ = new String[numberOfRegisters, 16];

                object[] comboBoxVal_ = new object[numberOfRegisters];

                for (int i = 0; i < numberOfRegisters; i++)
                {
                    textBoxLabel_[i] = textBoxLabel[i].Text;
                    textBoxVal_[i] = textBoxRegister[i].Text;
                    textBoxValue_[i] = textBoxValue[i].Text;

                    comboBoxVal_[i] = comboBox[i].SelectedItem;

                    for (int a = 0; a < 16; a++)
                    {
                        if(i == row)
                        {
                            labelBitRegisters_[i, a] = textBoxBitsLabel[a].Text;
                        }
                        else
                        {
                            labelBitRegisters_[i, a] = labelBitRegisters[i, a];
                        }
                    }
                }

                textBoxLabel_[row] = textBoxLabel_A.Text;
                textBoxVal_[row] = textBoxVal_A.Text;
                textBoxValue_[row] = textBoxValue_A.Text;

                comboBoxVal_[row] = comboBoxVal_A.SelectedItem;


                config.textBoxLabel_ = textBoxLabel_;
                config.textBoxVal_ = textBoxVal_;
                config.textBoxValue_ = textBoxValue_;
                config.comboBoxVal_ = comboBoxVal_;
                config.labelBitRegisters_ = labelBitRegisters_;

                string file_content = JsonConvert.SerializeObject(config);

                File.WriteAllText("Configurazione_form_2.json", file_content);
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }*/


        // Classe per caricare dati dal file di configurazione json
        public class SAVE_Form
        {
            public String textBoxModBusAddress_ { get; set; }
            public String textBoxHoldingOffset_ { get; set; }
            public object comboBoxHoldingOffset_ { get; set; }

            // Variabili array 
            public String[] textBoxLabel_ { get; set; }
            public String[] textBoxVal_ { get; set; }
            public String[] textBoxValue_ { get; set; }

            public object[] comboBoxVal_ { get; set; }

            public bool checkBoxOffBit_ { get; set; }
            public bool registriBloccati_ { get; set; }
            public bool disattivaComandiPictureBox_ { get; set; }
        }

        // Classe per caricare dati dal file di configurazione json
        public class SAVE_Form3
        {
            public String[] labelBitRegisters_ { get; set; }    // ToolTips bit
        }

        //-------------------------------------------------------------------------
        //---------------------Funzioni generali per elementi----------------------
        //-------------------------------------------------------------------------

        // COnverte un array di 16 pictureBox in un UInt16
        private UInt16 pictureToInt()
        {
            UInt16 val = 0;

            for (int i = 0; i < 16; i++)
            {
                if (pictureBixRegisterBits[i].Background != Brushes.LightGray)
                {
                    val += (UInt16)((UInt16)(1) << i);
                }
            }

            return val;
        }

        //Converte un valore UInt16 in una rappresentazione tramite un array di 16 pictureBox
        private void intToPicture(int value)
        {
            for (int i = 0; i < 16; i++)
            {
                if ((value & (1 << i)) > 0)
                {
                    pictureBixRegisterBits[i].Background = Brushes.Orange;
                }
                else
                {
                    pictureBixRegisterBits[i].Background = Brushes.LightGray;
                }
            }
        }

        //-------------------------------------------------------------------------
        //-------------------------Lettura regitro da PLC--------------------------
        //-------------------------------------------------------------------------


        private void buttonRead_A_Click(object sender, RoutedEventArgs e)
        {
            pictureBoxBusy.Background = Brushes.Yellow;

            UInt16 address = (UInt16)(P.uint_parser(textBoxVal_A, comboBoxVal_A) + P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset));

            UInt16 val = UInt16.Parse(ModBus.readHoldingRegister_03(byte.Parse(textBoxModBusAddress.Text), address, 1)[0]);

            intToPicture(val);
            textBoxValue_A.Text = val.ToString();

            pictureBoxBusy.Background = Brushes.LightGray;
        }

        //-------------------------------------------------------------------------
        //--------------------------Reset regitro su PLC--------------------------
        //-------------------------------------------------------------------------

        private void buttonReset_A_Click(object sender, RoutedEventArgs e)
        {
            pictureBoxBusy.Background = Brushes.Yellow;

            textBoxValue_A.Text = "0";

            for (int i = 0; i < 16; i++)
            {
                pictureBixRegisterBits[i].Background = Brushes.LightGray;
            }

            UInt16 address = (UInt16)(P.uint_parser(textBoxVal_A, comboBoxVal_A) + P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset));

            ModBus.presetSingleRegister_06(byte.Parse(textBoxModBusAddress.Text), address, 0);

            pictureBoxBusy.Background = Brushes.LightGray;
        }

        //-------------------------------------------------------------------------
        //----------------------Funzione click su picturebox-----------------------
        //-------------------------------------------------------------------------

        private void picture(int bit, bool force)
        {
            if (!(bool)sim_Auge.checkBoxDisablePictureBoxCommands.IsChecked || force)
            {
                pictureBoxBusy.Background = Brushes.Yellow;
                DoEvents();


                if (pictureBixRegisterBits[bit].Background == Brushes.LightGray)
                {
                    pictureBixRegisterBits[bit].Background = Brushes.Orange;
                    DoEvents();

                    if (sim_Auge != null)
                        sim_Auge.setPictureBoxFromOtherThread(row, bit, true, true);
                }
                else
                {
                    pictureBixRegisterBits[bit].Background = Brushes.LightGray;
                    DoEvents();

                    if (sim_Auge != null)
                        sim_Auge.setPictureBoxFromOtherThread(row, bit, false, true);
                }

                UInt16 val = pictureToInt();
                UInt16 address = (UInt16)(P.uint_parser(textBoxVal_A, comboBoxVal_A) + P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset));

                textBoxValue_A.Text = val.ToString();

                ModBus.presetSingleRegister_06(byte.Parse(textBoxModBusAddress.Text), address, val);

                pictureBoxBusy.Background = Brushes.LightGray;
                DoEvents();

                if (sim_Auge != null)
                    sim_Auge.setPictureBoxBusy(false);
            }
        }

        private void pictureBox_15_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(15, false);
        }

        private void pictureBox_14_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(14, false);
        }

        private void pictureBox_13_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(13, false);
        }

        private void pictureBox_12_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(12, false);
        }

        private void pictureBox_11_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(11, false);
        }

        private void pictureBox_10_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(10, false);
        }

        private void pictureBox_9_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(9, false);
        }

        private void pictureBox_8_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(8, false);
        }

        private void pictureBox_7_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, false);
        }

        private void pictureBox_6_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, false);
        }

        private void pictureBox_5_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, false);
        }

        private void pictureBox_4_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, false);
        }

        private void pictureBox_3_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, false);
        }

        private void pictureBox_2_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, false);
        }

        private void pictureBox_1_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, false);
        }

        private void pictureBox_0_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, false);
        }

        public void command(int bit)
        {
            picture(bit, true);
            Thread.Sleep(200);
            picture(bit, true);
        }

        private void buttonCommand_15_Click(object sender, RoutedEventArgs e)
        {
            command(15);
        }

        private void buttonCommand_14_Click(object sender, RoutedEventArgs e)
        {
            command(14);
        }

        private void buttonCommand_13_Click(object sender, RoutedEventArgs e)
        {
            command(13);
        }

        private void buttonCommand_12_Click(object sender, RoutedEventArgs e)
        {
            command(12);
        }

        private void buttonCommand_11_Click(object sender, RoutedEventArgs e)
        {
            command(11);
        }

        private void buttonCommand_10_Click(object sender, RoutedEventArgs e)
        {
            command(10);
        }

        private void buttonCommand_9_Click(object sender, RoutedEventArgs e)
        {
            command(9);
        }

        private void buttonCommand_8_Click(object sender, RoutedEventArgs e)
        {
            command(8);
        }

        private void buttonCommand_7_Click(object sender, RoutedEventArgs e)
        {
            command(7);
        }

        private void buttonCommand_6_Click(object sender, RoutedEventArgs e)
        {
            command(6);
        }

        private void buttonCommand_5_Click(object sender, RoutedEventArgs e)
        {
            command(5);
        }

        private void buttonCommand_4_Click(object sender, RoutedEventArgs e)
        {
            command(4);
        }

        private void buttonCommand_3_Click(object sender, RoutedEventArgs e)
        {
            command(3);
        }

        private void buttonCommand_2_Click(object sender, RoutedEventArgs e)
        {
            command(2);
        }

        private void buttonCommand_1_Click(object sender, RoutedEventArgs e)
        {
            command(1);
        }

        private void buttonCommand_0_Click(object sender, RoutedEventArgs e)
        {
            command(0);
        }

        private void salvaConfigurazioneToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione_3();
        }

        private void sbloccaRegistriToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            comboBoxHoldingOffset.IsEnabled = true;
            textBoxHoldingOffset.IsEnabled = true;
            comboBoxVal_A.IsEnabled = true;
            textBoxVal_A.IsEnabled = true;
            textBoxValue_A.IsEnabled = true;
        }

        private void buttonUp_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione_3();

            row++;

            if (row > 15)
                row = 0;

            carica_configurazione_3();
        }

        private void buttonDown_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione_3();

            row--;

            if (row < 0)
                row = 15;

            carica_configurazione_3();
        }

        private void sbloccaTextBoxToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                textBoxBitsLabel[i].IsEnabled = !textBoxBitsLabel[i].IsEnabled;
            }
        }

        // Funzione equivalnete alla vecchia Application.DoEvents()
        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
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
