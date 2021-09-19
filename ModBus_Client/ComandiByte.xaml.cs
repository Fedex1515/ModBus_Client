﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using ModBusMaster_Chicco;

using System.IO;

using Raccolta_funzioni_parser;

//Libreria JSON
using System.Web.Script.Serialization;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for ComandiByte.xaml
    /// </summary>
    public partial class ComandiByte : Window
    {
        ModBus_Chicco ModBus;
        MainWindow modBus_Client;

        int page = 0;
        static int numberOfPages = 4;  // Numero di registri comandabili dal form

        String[] S_textBoxLabel = new String[16 * numberOfPages];
        String[] S_comboBoxHoldingAddress = new String[16 * numberOfPages];
        String[] S_textBoxHoldingAddress = new String[16 * numberOfPages];
        String[] S_comboBoxHoldingValue = new String[16 * numberOfPages];
        String[] S_textBoxHoldingValue = new String[16 * numberOfPages];


        TextBox[] textBoxLabel = new TextBox[16];
        TextBox[] textBoxHoldingAddress = new TextBox[16];
        TextBox[] textBoxHoldingValue = new TextBox[16];
        ComboBox[] comboBoxHoldingAddress = new ComboBox[16];
        ComboBox[] comboBoxHoldingValue = new ComboBox[16];

        Button[] buttonRead = new Button[16];
        Button[] buttonReset = new Button[16];

        Parser P = new Parser();

        String pathToConfiguration;

        public ComandiByte(ModBus_Chicco ModBus_, MainWindow modBus_Client_)
        {
            InitializeComponent();

            // Creo evento di chiusura del form
            this.Closing += Sim_Form_cs_Closing;

            ModBus = ModBus_;
            modBus_Client = modBus_Client_;
            pathToConfiguration = modBus_Client.pathToConfiguration;

            textBoxModBusAddress.Text = modBus_Client.textBoxModbusAddress.Text;

            textBoxHoldingOffset.Text = modBus_Client.textBoxHoldingOffset.Text;
            comboBoxHoldingOffset.SelectedIndex = modBus_Client.comboBoxHoldingOffset.SelectedIndex;

            // Assegnazione array elementi grafici
            textBoxLabel[0] = textBoxLabel_A;
            textBoxLabel[1] = textBoxLabel_B;
            textBoxLabel[2] = textBoxLabel_C;
            textBoxLabel[3] = textBoxLabel_D;
            textBoxLabel[4] = textBoxLabel_E;
            textBoxLabel[5] = textBoxLabel_F;
            textBoxLabel[6] = textBoxLabel_G;
            textBoxLabel[7] = textBoxLabel_H;
            textBoxLabel[8] = textBoxLabel_I;
            textBoxLabel[9] = textBoxLabel_J;
            textBoxLabel[10] = textBoxLabel_K;
            textBoxLabel[11] = textBoxLabel_L;
            textBoxLabel[12] = textBoxLabel_M;
            textBoxLabel[13] = textBoxLabel_N;
            textBoxLabel[14] = textBoxLabel_O;
            textBoxLabel[15] = textBoxLabel_P;

            textBoxHoldingAddress[0] = textBoxHoldingAddress_A;
            textBoxHoldingAddress[1] = textBoxHoldingAddress_B;
            textBoxHoldingAddress[2] = textBoxHoldingAddress_C;
            textBoxHoldingAddress[3] = textBoxHoldingAddress_D;
            textBoxHoldingAddress[4] = textBoxHoldingAddress_E;
            textBoxHoldingAddress[5] = textBoxHoldingAddress_F;
            textBoxHoldingAddress[6] = textBoxHoldingAddress_G;
            textBoxHoldingAddress[7] = textBoxHoldingAddress_H;
            textBoxHoldingAddress[8] = textBoxHoldingAddress_I;
            textBoxHoldingAddress[9] = textBoxHoldingAddress_J;
            textBoxHoldingAddress[10] = textBoxHoldingAddress_K;
            textBoxHoldingAddress[11] = textBoxHoldingAddress_L;
            textBoxHoldingAddress[12] = textBoxHoldingAddress_M;
            textBoxHoldingAddress[13] = textBoxHoldingAddress_N;
            textBoxHoldingAddress[14] = textBoxHoldingAddress_O;
            textBoxHoldingAddress[15] = textBoxHoldingAddress_P;

            textBoxHoldingValue[0] = textBoxHoldingValue_A;
            textBoxHoldingValue[1] = textBoxHoldingValue_B;
            textBoxHoldingValue[2] = textBoxHoldingValue_C;
            textBoxHoldingValue[3] = textBoxHoldingValue_D;
            textBoxHoldingValue[4] = textBoxHoldingValue_E;
            textBoxHoldingValue[5] = textBoxHoldingValue_F;
            textBoxHoldingValue[6] = textBoxHoldingValue_G;
            textBoxHoldingValue[7] = textBoxHoldingValue_H;
            textBoxHoldingValue[8] = textBoxHoldingValue_I;
            textBoxHoldingValue[9] = textBoxHoldingValue_J;
            textBoxHoldingValue[10] = textBoxHoldingValue_K;
            textBoxHoldingValue[11] = textBoxHoldingValue_L;
            textBoxHoldingValue[12] = textBoxHoldingValue_M;
            textBoxHoldingValue[13] = textBoxHoldingValue_N;
            textBoxHoldingValue[14] = textBoxHoldingValue_O;
            textBoxHoldingValue[15] = textBoxHoldingValue_P;

            comboBoxHoldingAddress[0] = comboBoxHoldingAddress_A;
            comboBoxHoldingAddress[1] = comboBoxHoldingAddress_B;
            comboBoxHoldingAddress[2] = comboBoxHoldingAddress_C;
            comboBoxHoldingAddress[3] = comboBoxHoldingAddress_D;
            comboBoxHoldingAddress[4] = comboBoxHoldingAddress_E;
            comboBoxHoldingAddress[5] = comboBoxHoldingAddress_F;
            comboBoxHoldingAddress[6] = comboBoxHoldingAddress_G;
            comboBoxHoldingAddress[7] = comboBoxHoldingAddress_H;
            comboBoxHoldingAddress[8] = comboBoxHoldingAddress_I;
            comboBoxHoldingAddress[9] = comboBoxHoldingAddress_J;
            comboBoxHoldingAddress[10] = comboBoxHoldingAddress_K;
            comboBoxHoldingAddress[11] = comboBoxHoldingAddress_L;
            comboBoxHoldingAddress[12] = comboBoxHoldingAddress_M;
            comboBoxHoldingAddress[13] = comboBoxHoldingAddress_N;
            comboBoxHoldingAddress[14] = comboBoxHoldingAddress_O;
            comboBoxHoldingAddress[15] = comboBoxHoldingAddress_P;

            comboBoxHoldingValue[0] = comboBoxHoldingValue_A;
            comboBoxHoldingValue[1] = comboBoxHoldingValue_B;
            comboBoxHoldingValue[2] = comboBoxHoldingValue_C;
            comboBoxHoldingValue[3] = comboBoxHoldingValue_D;
            comboBoxHoldingValue[4] = comboBoxHoldingValue_E;
            comboBoxHoldingValue[5] = comboBoxHoldingValue_F;
            comboBoxHoldingValue[6] = comboBoxHoldingValue_G;
            comboBoxHoldingValue[7] = comboBoxHoldingValue_H;
            comboBoxHoldingValue[8] = comboBoxHoldingValue_I;
            comboBoxHoldingValue[9] = comboBoxHoldingValue_J;
            comboBoxHoldingValue[10] = comboBoxHoldingValue_K;
            comboBoxHoldingValue[11] = comboBoxHoldingValue_L;
            comboBoxHoldingValue[12] = comboBoxHoldingValue_M;
            comboBoxHoldingValue[13] = comboBoxHoldingValue_N;
            comboBoxHoldingValue[14] = comboBoxHoldingValue_O;
            comboBoxHoldingValue[15] = comboBoxHoldingValue_P;

            buttonRead[0] = buttonRead_A;
            buttonRead[1] = buttonRead_B;
            buttonRead[2] = buttonRead_C;
            buttonRead[3] = buttonRead_D;
            buttonRead[4] = buttonRead_E;
            buttonRead[5] = buttonRead_F;
            buttonRead[6] = buttonRead_G;
            buttonRead[7] = buttonRead_H;
            buttonRead[8] = buttonRead_I;
            buttonRead[9] = buttonRead_J;
            buttonRead[10] = buttonRead_K;
            buttonRead[11] = buttonRead_L;
            buttonRead[12] = buttonRead_M;
            buttonRead[13] = buttonRead_N;
            buttonRead[14] = buttonRead_O;
            buttonRead[15] = buttonRead_P;

            buttonReset[0] = buttonReset_A;
            buttonReset[1] = buttonReset_B;
            buttonReset[2] = buttonReset_C;
            buttonReset[3] = buttonReset_D;
            buttonReset[4] = buttonReset_E;
            buttonReset[5] = buttonReset_F;
            buttonReset[6] = buttonReset_G;
            buttonReset[7] = buttonReset_H;
            buttonReset[8] = buttonReset_I;
            buttonReset[9] = buttonReset_J;
            buttonReset[10] = buttonReset_K;
            buttonReset[11] = buttonReset_L;
            buttonReset[12] = buttonReset_M;
            buttonReset[13] = buttonReset_N;
            buttonReset[14] = buttonReset_O;
            buttonReset[15] = buttonReset_P;

            // Grafica di default
            comboBoxHoldingOffset.SelectedIndex = 0;

            for (int i = 0; i < 16; i++)
            {
                comboBoxHoldingAddress[i].SelectedIndex = 0;
                comboBoxHoldingValue[i].SelectedIndex = 0;

                // Nascondo le caselle pari
                if (i % 2 == 1)
                {
                    buttonRead[i].Visibility = Visibility.Hidden;
                    buttonReset[i].Visibility = Visibility.Hidden;
                }
            }

            carica_configurazione_4();

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
        }

        public void Sim_Form_cs_Closing(object sender, EventArgs e)
        {
            salva_configurazione_4();
        }

        private void carica_configurazione_4()
        {
            //Caricamento valori ultima sessione di questo form
            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/ComandiByte.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                SAVE_Form4 config = jss.Deserialize<SAVE_Form4>(file_content);

                S_textBoxLabel = config.textBoxLabel_;
                S_comboBoxHoldingAddress = config.comboBoxHoldingAddress_;
                S_textBoxHoldingAddress = config.textBoxHoldingAddress_;
                S_comboBoxHoldingValue = config.comboBoxHoldingValue_;
                S_textBoxHoldingValue = config.textBoxHoldingValue_;

                for (int i = 0; i < 16; i++)
                {
                    textBoxLabel[i].Text = S_textBoxLabel[i];
                    comboBoxHoldingAddress[i].SelectedIndex = S_comboBoxHoldingAddress[i] == "HEX" ? 1 : 0;
                    textBoxHoldingAddress[i].Text = S_textBoxHoldingAddress[i];
                    comboBoxHoldingValue[i].SelectedIndex = S_comboBoxHoldingValue[i] == "HEX" ? 1 : 0;
                    textBoxHoldingValue[i].Text = S_textBoxHoldingValue[i];
                }

                if (config.registriBloccati_)
                {
                    for (int i = 0; i < 16; i++)
                    {
                        textBoxLabel[i].IsEnabled = false;
                        comboBoxHoldingAddress[i].IsEnabled = false;
                        textBoxHoldingAddress[i].IsEnabled = false;
                        comboBoxHoldingValue[i].IsEnabled = false;
                        textBoxHoldingValue[i].IsEnabled = false;
                    }

                    textBoxModBusAddress.IsEnabled = textBoxLabel[0].IsEnabled;
                    comboBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
                    textBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }


        public void salva_configurazione_4()
        {
            salvaGrafica();

            //Salvataggio valori ultima sessione
            try
            {
                SAVE_Form4 config = new SAVE_Form4();

                config.textBoxLabel_ = S_textBoxLabel;
                config.comboBoxHoldingAddress_ = S_comboBoxHoldingAddress;
                config.textBoxHoldingAddress_ = S_textBoxHoldingAddress;
                config.comboBoxHoldingValue_ = S_comboBoxHoldingValue;
                config.textBoxHoldingValue_ = S_textBoxHoldingValue;

                config.registriBloccati_ = !textBoxLabel[0].IsEnabled;

                JavaScriptSerializer jss = new JavaScriptSerializer();
                string file_content = jss.Serialize(config);

                File.WriteAllText("Json/" + pathToConfiguration + "/ComandiByte.json", file_content);

            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        // Classe per caricare dati dal file di configurazione json
        public class SAVE_Form4
        {
            public String[] textBoxLabel_ { get; set; }
            public String[] comboBoxHoldingAddress_ { get; set; }
            public String[] textBoxHoldingAddress_ { get; set; }
            public String[] comboBoxHoldingValue_ { get; set; }
            public String[] textBoxHoldingValue_ { get; set; }

            public bool registriBloccati_ { get; set; }
        }

        private void buttonUp_Click(object sender, RoutedEventArgs e)
        {
            salvaGrafica();

            page++;

            if (page > 3)
                page = 0;

            // Cancello i colori delle textBox
            for (int i = 0; i < 16; i++)
            {
                textBoxHoldingValue[i].Background = Brushes.White;
            }

            aggiornaGrafica();
        }

        private void buttonDown_Click(object sender, RoutedEventArgs e)
        {
            salvaGrafica();

            page--;

            if (page < 0)
                page = 3;

            // Cancello i colori delle textBox
            for (int i = 0; i < 16; i++)
                textBoxHoldingValue[i].Background = Brushes.White;

            aggiornaGrafica();
        }

        private void aggiornaGrafica()
        {
            for (int i = 0; i < 16; i++)
            {
                textBoxLabel[i].Text = S_textBoxLabel[i + page * 16];
                comboBoxHoldingAddress[i].SelectedIndex = S_comboBoxHoldingAddress[i + page * 16] == "HEX" ? 1 : 0;
                textBoxHoldingAddress[i].Text = S_textBoxHoldingAddress[i + page * 16];
                comboBoxHoldingValue[i].SelectedIndex = S_comboBoxHoldingValue[i + page * 16] == "HEX" ? 1 : 0;
                textBoxHoldingValue[i].Text = S_textBoxHoldingValue[i + page * 16];
            }
        }

        private void salvaGrafica()
        {
            for (int i = 0; i < 16; i++)
            {
                S_textBoxLabel[i + page * 16] = textBoxLabel[i].Text;
                S_comboBoxHoldingAddress[i + page * 16] = comboBoxHoldingAddress[i].SelectedItem.ToString().Split(' ')[1];
                S_textBoxHoldingAddress[i + page * 16] = textBoxHoldingAddress[i].Text;
                S_comboBoxHoldingValue[i + page * 16] = comboBoxHoldingValue[i].SelectedItem.ToString().Split(' ')[1];
                S_textBoxHoldingValue[i + page * 16] = textBoxHoldingValue[i].Text;
            }
        }

        private void salvaConfigurazioneToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione_4();
        }

        public void read(int row)
        {
            pictureBoxBusy.Background = Brushes.Yellow;
            DoEvents();

            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingAddress[row * 2], comboBoxHoldingAddress[row * 2]);

                string[] response = ModBus.readHoldingRegister_03(byte.Parse(textBoxModBusAddress.Text), address_start, 1);

                uint[] response_ = new uint[response.Length];

                for (int i = 0; i < response.Length; i++)
                    response_[i] = uint.Parse(response[i]);

                if (comboBoxHoldingValue[row * 2].SelectedItem.ToString() == "DEC")
                {
                    // Visualizzazione in decimale
                    textBoxHoldingValue[row * 2 + 1].Text = (response_[0] & 0x00FF).ToString();     // LSB
                    textBoxHoldingValue[row * 2].Text = ((response_[0] & 0xFF00) >> 8).ToString();    // MSB
                }
                else
                {
                    // Visualizzazione in hex
                    textBoxHoldingValue[row * 2 + 1].Text = (response_[0] & 0x00FF).ToString("X").PadLeft(2, '0'); ;    // LSB
                    textBoxHoldingValue[row * 2].Text = ((response_[0] & 0xFF00) >> 8).ToString("X").PadLeft(2, '0'); ; // MSB
                }

                textBoxHoldingValue[row * 2 + 1].Background = Brushes.LightBlue;
                textBoxHoldingValue[row * 2].Background = Brushes.LightBlue;
            }
            catch { }

            pictureBoxBusy.Background = Brushes.LightGray;
            DoEvents();
        }

        public void write(int row)
        {
            pictureBoxBusy.Background = Brushes.Yellow;
            DoEvents();

            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingAddress[row], comboBoxHoldingAddress[row]);
                uint value = (P.uint_parser(textBoxHoldingValue[row], comboBoxHoldingValue[row]) << 8) + P.uint_parser(textBoxHoldingValue[row + 1], comboBoxHoldingValue[row + 1]);

                //DEBUG
                Console.WriteLine("Value: " + value.ToString());

                if (value < 65536 && value >= 0)
                {

                    if (ModBus.presetSingleRegister_06(byte.Parse(textBoxModBusAddress.Text), address_start, value))
                    {
                        textBoxHoldingValue[row].Background = Brushes.LightGreen;
                        textBoxHoldingValue[row + 1].Background = Brushes.LightGreen;
                    }
                    else
                    {
                        textBoxHoldingValue[row].Background = Brushes.PaleVioletRed;
                        textBoxHoldingValue[row + 1].Background = Brushes.PaleVioletRed;
                        MessageBox.Show("Errore nel settaggio del registro", "Alert");
                    }
                }
                else
                {
                    Console.WriteLine("Valore inserito non valido: " + value.ToString());
                    MessageBox.Show("Valore inserito non valido: " + value.ToString(), "Alert");
                }
            }
            catch { }


            pictureBoxBusy.Background = Brushes.LightGray;
            DoEvents();
        }


        private void buttonWrite_A_Click(object sender, RoutedEventArgs e)
        {
            write(0);
        }


        private void buttonWrite_B_Click(object sender, RoutedEventArgs e)
        {
            write(0);
        }

        private void buttonWrite_C_Click(object sender, RoutedEventArgs e)
        {
            write(2);
        }

        private void buttonWrite_D_Click(object sender, RoutedEventArgs e)
        {
            write(2);
        }

        private void buttonWrite_E_Click(object sender, RoutedEventArgs e)
        {
            write(4);
        }

        private void buttonWrite_F_Click(object sender, RoutedEventArgs e)
        {
            write(4);
        }

        private void buttonWrite_G_Click(object sender, RoutedEventArgs e)
        {
            write(6);
        }

        private void buttonWrite_H_Click(object sender, RoutedEventArgs e)
        {
            write(6);
        }

        private void buttonWrite_I_Click(object sender, RoutedEventArgs e)
        {
            write(8);
        }

        private void buttonWrite_J_Click(object sender, RoutedEventArgs e)
        {
            write(8);
        }

        private void buttonWrite_K_Click(object sender, RoutedEventArgs e)
        {
            write(10);
        }

        private void buttonWrite_L_Click(object sender, RoutedEventArgs e)
        {
            write(10);
        }

        private void buttonWrite_M_Click(object sender, RoutedEventArgs e)
        {
            write(12);
        }

        private void buttonWrite_N_Click(object sender, RoutedEventArgs e)
        {
            write(12);
        }

        private void buttonWrite_O_Click(object sender, RoutedEventArgs e)
        {
            write(14);
        }

        private void buttonWrite_P_Click(object sender, RoutedEventArgs e)
        {
            write(14);
        }

        private void buttonRead_A_Click(object sender, RoutedEventArgs e)
        {
            read(0);
        }

        private void buttonRead_B_Click(object sender, RoutedEventArgs e)
        {
            read(0);
        }

        private void buttonRead_C_Click(object sender, RoutedEventArgs e)
        {
            read(1);
        }

        private void buttonRead_D_Click(object sender, RoutedEventArgs e)
        {
            read(1);
        }

        private void buttonRead_E_Click(object sender, RoutedEventArgs e)
        {
            read(2);
        }

        private void buttonRead_F_Click(object sender, RoutedEventArgs e)
        {
            read(2);
        }

        private void buttonRead_G_Click(object sender, RoutedEventArgs e)
        {
            read(3);
        }

        private void buttonRead_H_Click(object sender, RoutedEventArgs e)
        {
            read(3);
        }

        private void buttonRead_I_Click(object sender, RoutedEventArgs e)
        {
            read(4);
        }

        private void buttonRead_J_Click(object sender, RoutedEventArgs e)
        {
            read(4);
        }

        private void buttonRead_K_Click(object sender, RoutedEventArgs e)
        {
            read(5);
        }

        private void buttonRead_L_Click(object sender, RoutedEventArgs e)
        {
            read(5);
        }

        private void buttonRead_M_Click(object sender, RoutedEventArgs e)
        {
            read(6);
        }

        private void buttonRead_N_Click(object sender, RoutedEventArgs e)
        {
            read(6);
        }

        private void buttonRead_O_Click(object sender, RoutedEventArgs e)
        {
            read(7);
        }

        private void buttonRead_P_Click(object sender, RoutedEventArgs e)
        {
            read(7);
        }

        private void buttonReadAll_Click(object sender, RoutedEventArgs e)
        {
            pictureBoxBusy.Background = Brushes.Yellow;

            for (int i = 0; i < 8; i++)
            {
                read(i);
            }

            pictureBoxBusy.Background = Brushes.LightGray;
        }

        private void buttonResetAll_Click(object sender, RoutedEventArgs e)
        {
            pictureBoxBusy.Background = Brushes.Yellow;

            for (int i = 0; i < 16; i++)
            {
                write(i);
            }

            pictureBoxBusy.Background = Brushes.LightGray;
        }

        private void bloccasbloccaRegistriToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                textBoxLabel[i].IsEnabled = !textBoxLabel[i].IsEnabled;
                comboBoxHoldingAddress[i].IsEnabled = textBoxLabel[i].IsEnabled;
                textBoxHoldingAddress[i].IsEnabled = textBoxLabel[i].IsEnabled;
                comboBoxHoldingValue[i].IsEnabled = textBoxLabel[i].IsEnabled;
                textBoxHoldingValue[i].IsEnabled = textBoxLabel[i].IsEnabled;
            }

            textBoxModBusAddress.IsEnabled = textBoxLabel[0].IsEnabled;
            comboBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
            textBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
        }

        private void MenuItemSalvaConfigurazione_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione_4();
        }

        private void MenuItemBloccaSbloccaTextBoxes_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                textBoxLabel[i].IsEnabled = !textBoxLabel[i].IsEnabled;
                comboBoxHoldingAddress[i].IsEnabled = textBoxLabel[i].IsEnabled;
                textBoxHoldingAddress[i].IsEnabled = textBoxLabel[i].IsEnabled;
                comboBoxHoldingValue[i].IsEnabled = textBoxLabel[i].IsEnabled;
                textBoxHoldingValue[i].IsEnabled = textBoxLabel[i].IsEnabled;
            }

            textBoxModBusAddress.IsEnabled = textBoxLabel[0].IsEnabled;
            comboBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
            textBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
        }

        // Funzione equivalnete alla vecchia Application.DoEvents()
        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
