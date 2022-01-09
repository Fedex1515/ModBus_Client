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

using ModBusMaster_Chicco;

using System.IO;

using Raccolta_funzioni_parser;

using Microsoft.Win32;

//Libreria JSON
using System.Web.Script.Serialization;

// Libreria lingue
using LanguageLib; // Libreria custom per caricare etichette in lingue differenti

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for ComandiBit.xaml
    /// </summary>
    public partial class ComandiBit : Window
    {
        ModBus_Chicco ModBus;
        MainWindow modBus_Client;
        Language lang; 

        static int numberOfRegisters = 16;  // Numero di registri comandabili dal form

        public TextBox[] textBoxRegister = new TextBox[numberOfRegisters];
        public TextBox[] textBoxLabel = new TextBox[numberOfRegisters];
        public TextBox[] textBoxValue = new TextBox[numberOfRegisters];
        ComboBox[] comboBox = new ComboBox[numberOfRegisters];
        Border[,] pictureBox = new Border[numberOfRegisters, 16];

        String[,] labelBitRegisters = new String[numberOfRegisters, 16];

        Parser P = new Parser();

        //ToolTips
        // Create the ToolTip and associate with the Form container.
        ToolTip toolTipHelp = new ToolTip();

        String pathToConfiguration;

        public ComandiBit(ModBus_Chicco ModBus_, MainWindow modBus_Client_)
        {
            InitializeComponent();

            lang = new Language(this);

            // Creo evento di chiusura del form
            this.Closing += Sim_Form_cs_Closing;

            ModBus = ModBus_;
            modBus_Client = modBus_Client_;
            pathToConfiguration = modBus_Client.pathToConfiguration;

            modBus_Client.toolSTripMenuEnable(false);

            textBoxModBusAddress.Text = modBus_Client.textBoxModbusAddress.Text;

            textBoxHoldingOffset.Text = modBus_Client.textBoxHoldingOffset.Text;
            comboBoxHoldingOffset.SelectedIndex = modBus_Client.comboBoxHoldingOffset.SelectedIndex;

            //Assegnazione textBox
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

            //Assegnazione textBox
            textBoxRegister[0] = textBoxVal_A;
            textBoxRegister[1] = textBoxVal_B;
            textBoxRegister[2] = textBoxVal_C;
            textBoxRegister[3] = textBoxVal_D;
            textBoxRegister[4] = textBoxVal_E;
            textBoxRegister[5] = textBoxVal_F;
            textBoxRegister[6] = textBoxVal_G;
            textBoxRegister[7] = textBoxVal_H;
            textBoxRegister[8] = textBoxVal_I;
            textBoxRegister[9] = textBoxVal_J;
            textBoxRegister[10] = textBoxVal_K;
            textBoxRegister[11] = textBoxVal_L;
            textBoxRegister[12] = textBoxVal_M;
            textBoxRegister[13] = textBoxVal_N;
            textBoxRegister[14] = textBoxVal_O;
            textBoxRegister[15] = textBoxVal_P;

            //Assegnazione textBox
            textBoxValue[0] = textBoxValue_A;
            textBoxValue[1] = textBoxValue_B;
            textBoxValue[2] = textBoxValue_C;
            textBoxValue[3] = textBoxValue_D;
            textBoxValue[4] = textBoxValue_E;
            textBoxValue[5] = textBoxValue_F;
            textBoxValue[6] = textBoxValue_G;
            textBoxValue[7] = textBoxValue_H;
            textBoxValue[8] = textBoxValue_I;
            textBoxValue[9] = textBoxValue_J;
            textBoxValue[10] = textBoxValue_K;
            textBoxValue[11] = textBoxValue_L;
            textBoxValue[12] = textBoxValue_M;
            textBoxValue[13] = textBoxValue_N;
            textBoxValue[14] = textBoxValue_O;
            textBoxValue[15] = textBoxValue_P;

            //Assegnazione comboBox
            comboBox[0] = comboBoxVal_A;
            comboBox[1] = comboBoxVal_B;
            comboBox[2] = comboBoxVal_C;
            comboBox[3] = comboBoxVal_D;
            comboBox[4] = comboBoxVal_E;
            comboBox[5] = comboBoxVal_F;
            comboBox[6] = comboBoxVal_G;
            comboBox[7] = comboBoxVal_H;
            comboBox[8] = comboBoxVal_I;
            comboBox[9] = comboBoxVal_J;
            comboBox[10] = comboBoxVal_K;
            comboBox[11] = comboBoxVal_L;
            comboBox[12] = comboBoxVal_M;
            comboBox[13] = comboBoxVal_N;
            comboBox[14] = comboBoxVal_O;
            comboBox[15] = comboBoxVal_P;

            //Assegnazione pictureBox

            //Riga 0
            pictureBox[0, 0] = pictureBox_00_A;
            pictureBox[0, 1] = pictureBox_01_A;
            pictureBox[0, 2] = pictureBox_02_A;
            pictureBox[0, 3] = pictureBox_03_A;
            pictureBox[0, 4] = pictureBox_04_A;
            pictureBox[0, 5] = pictureBox_05_A;
            pictureBox[0, 6] = pictureBox_06_A;
            pictureBox[0, 7] = pictureBox_07_A;
            pictureBox[0, 8] = pictureBox_08_A;
            pictureBox[0, 9] = pictureBox_09_A;
            pictureBox[0, 10] = pictureBox_10_A;
            pictureBox[0, 11] = pictureBox_11_A;
            pictureBox[0, 12] = pictureBox_12_A;
            pictureBox[0, 13] = pictureBox_13_A;
            pictureBox[0, 14] = pictureBox_14_A;
            pictureBox[0, 15] = pictureBox_15_A;

            //Riga 1
            pictureBox[1, 0] = pictureBox_00_B;
            pictureBox[1, 1] = pictureBox_01_B;
            pictureBox[1, 2] = pictureBox_02_B;
            pictureBox[1, 3] = pictureBox_03_B;
            pictureBox[1, 4] = pictureBox_04_B;
            pictureBox[1, 5] = pictureBox_05_B;
            pictureBox[1, 6] = pictureBox_06_B;
            pictureBox[1, 7] = pictureBox_07_B;
            pictureBox[1, 8] = pictureBox_08_B;
            pictureBox[1, 9] = pictureBox_09_B;
            pictureBox[1, 10] = pictureBox_10_B;
            pictureBox[1, 11] = pictureBox_11_B;
            pictureBox[1, 12] = pictureBox_12_B;
            pictureBox[1, 13] = pictureBox_13_B;
            pictureBox[1, 14] = pictureBox_14_B;
            pictureBox[1, 15] = pictureBox_15_B;

            //Riga 2
            pictureBox[2, 0] = pictureBox_00_C;
            pictureBox[2, 1] = pictureBox_01_C;
            pictureBox[2, 2] = pictureBox_02_C;
            pictureBox[2, 3] = pictureBox_03_C;
            pictureBox[2, 4] = pictureBox_04_C;
            pictureBox[2, 5] = pictureBox_05_C;
            pictureBox[2, 6] = pictureBox_06_C;
            pictureBox[2, 7] = pictureBox_07_C;
            pictureBox[2, 8] = pictureBox_08_C;
            pictureBox[2, 9] = pictureBox_09_C;
            pictureBox[2, 10] = pictureBox_10_C;
            pictureBox[2, 11] = pictureBox_11_C;
            pictureBox[2, 12] = pictureBox_12_C;
            pictureBox[2, 13] = pictureBox_13_C;
            pictureBox[2, 14] = pictureBox_14_C;
            pictureBox[2, 15] = pictureBox_15_C;

            //Riga 3
            pictureBox[3, 0] = pictureBox_00_D;
            pictureBox[3, 1] = pictureBox_01_D;
            pictureBox[3, 2] = pictureBox_02_D;
            pictureBox[3, 3] = pictureBox_03_D;
            pictureBox[3, 4] = pictureBox_04_D;
            pictureBox[3, 5] = pictureBox_05_D;
            pictureBox[3, 6] = pictureBox_06_D;
            pictureBox[3, 7] = pictureBox_07_D;
            pictureBox[3, 8] = pictureBox_08_D;
            pictureBox[3, 9] = pictureBox_09_D;
            pictureBox[3, 10] = pictureBox_10_D;
            pictureBox[3, 11] = pictureBox_11_D;
            pictureBox[3, 12] = pictureBox_12_D;
            pictureBox[3, 13] = pictureBox_13_D;
            pictureBox[3, 14] = pictureBox_14_D;
            pictureBox[3, 15] = pictureBox_15_D;

            //Riga 4
            pictureBox[4, 0] = pictureBox_00_E;
            pictureBox[4, 1] = pictureBox_01_E;
            pictureBox[4, 2] = pictureBox_02_E;
            pictureBox[4, 3] = pictureBox_03_E;
            pictureBox[4, 4] = pictureBox_04_E;
            pictureBox[4, 5] = pictureBox_05_E;
            pictureBox[4, 6] = pictureBox_06_E;
            pictureBox[4, 7] = pictureBox_07_E;
            pictureBox[4, 8] = pictureBox_08_E;
            pictureBox[4, 9] = pictureBox_09_E;
            pictureBox[4, 10] = pictureBox_10_E;
            pictureBox[4, 11] = pictureBox_11_E;
            pictureBox[4, 12] = pictureBox_12_E;
            pictureBox[4, 13] = pictureBox_13_E;
            pictureBox[4, 14] = pictureBox_14_E;
            pictureBox[4, 15] = pictureBox_15_E;

            //Riga 5
            pictureBox[5, 0] = pictureBox_00_F;
            pictureBox[5, 1] = pictureBox_01_F;
            pictureBox[5, 2] = pictureBox_02_F;
            pictureBox[5, 3] = pictureBox_03_F;
            pictureBox[5, 4] = pictureBox_04_F;
            pictureBox[5, 5] = pictureBox_05_F;
            pictureBox[5, 6] = pictureBox_06_F;
            pictureBox[5, 7] = pictureBox_07_F;
            pictureBox[5, 8] = pictureBox_08_F;
            pictureBox[5, 9] = pictureBox_09_F;
            pictureBox[5, 10] = pictureBox_10_F;
            pictureBox[5, 11] = pictureBox_11_F;
            pictureBox[5, 12] = pictureBox_12_F;
            pictureBox[5, 13] = pictureBox_13_F;
            pictureBox[5, 14] = pictureBox_14_F;
            pictureBox[5, 15] = pictureBox_15_F;

            //Riga 6
            pictureBox[6, 0] = pictureBox_00_G;
            pictureBox[6, 1] = pictureBox_01_G;
            pictureBox[6, 2] = pictureBox_02_G;
            pictureBox[6, 3] = pictureBox_03_G;
            pictureBox[6, 4] = pictureBox_04_G;
            pictureBox[6, 5] = pictureBox_05_G;
            pictureBox[6, 6] = pictureBox_06_G;
            pictureBox[6, 7] = pictureBox_07_G;
            pictureBox[6, 8] = pictureBox_08_G;
            pictureBox[6, 9] = pictureBox_09_G;
            pictureBox[6, 10] = pictureBox_10_G;
            pictureBox[6, 11] = pictureBox_11_G;
            pictureBox[6, 12] = pictureBox_12_G;
            pictureBox[6, 13] = pictureBox_13_G;
            pictureBox[6, 14] = pictureBox_14_G;
            pictureBox[6, 15] = pictureBox_15_G;

            //Riga 7
            pictureBox[7, 0] = pictureBox_00_H;
            pictureBox[7, 1] = pictureBox_01_H;
            pictureBox[7, 2] = pictureBox_02_H;
            pictureBox[7, 3] = pictureBox_03_H;
            pictureBox[7, 4] = pictureBox_04_H;
            pictureBox[7, 5] = pictureBox_05_H;
            pictureBox[7, 6] = pictureBox_06_H;
            pictureBox[7, 7] = pictureBox_07_H;
            pictureBox[7, 8] = pictureBox_08_H;
            pictureBox[7, 9] = pictureBox_09_H;
            pictureBox[7, 10] = pictureBox_10_H;
            pictureBox[7, 11] = pictureBox_11_H;
            pictureBox[7, 12] = pictureBox_12_H;
            pictureBox[7, 13] = pictureBox_13_H;
            pictureBox[7, 14] = pictureBox_14_H;
            pictureBox[7, 15] = pictureBox_15_H;

            //Riga 8
            pictureBox[8, 0] = pictureBox_00_I;
            pictureBox[8, 1] = pictureBox_01_I;
            pictureBox[8, 2] = pictureBox_02_I;
            pictureBox[8, 3] = pictureBox_03_I;
            pictureBox[8, 4] = pictureBox_04_I;
            pictureBox[8, 5] = pictureBox_05_I;
            pictureBox[8, 6] = pictureBox_06_I;
            pictureBox[8, 7] = pictureBox_07_I;
            pictureBox[8, 8] = pictureBox_08_I;
            pictureBox[8, 9] = pictureBox_09_I;
            pictureBox[8, 10] = pictureBox_10_I;
            pictureBox[8, 11] = pictureBox_11_I;
            pictureBox[8, 12] = pictureBox_12_I;
            pictureBox[8, 13] = pictureBox_13_I;
            pictureBox[8, 14] = pictureBox_14_I;
            pictureBox[8, 15] = pictureBox_15_I;

            //Riga 9
            pictureBox[9, 0] = pictureBox_00_J;
            pictureBox[9, 1] = pictureBox_01_J;
            pictureBox[9, 2] = pictureBox_02_J;
            pictureBox[9, 3] = pictureBox_03_J;
            pictureBox[9, 4] = pictureBox_04_J;
            pictureBox[9, 5] = pictureBox_05_J;
            pictureBox[9, 6] = pictureBox_06_J;
            pictureBox[9, 7] = pictureBox_07_J;
            pictureBox[9, 8] = pictureBox_08_J;
            pictureBox[9, 9] = pictureBox_09_J;
            pictureBox[9, 10] = pictureBox_10_J;
            pictureBox[9, 11] = pictureBox_11_J;
            pictureBox[9, 12] = pictureBox_12_J;
            pictureBox[9, 13] = pictureBox_13_J;
            pictureBox[9, 14] = pictureBox_14_J;
            pictureBox[9, 15] = pictureBox_15_J;

            //Riga 10
            pictureBox[10, 0] = pictureBox_00_K;
            pictureBox[10, 1] = pictureBox_01_K;
            pictureBox[10, 2] = pictureBox_02_K;
            pictureBox[10, 3] = pictureBox_03_K;
            pictureBox[10, 4] = pictureBox_04_K;
            pictureBox[10, 5] = pictureBox_05_K;
            pictureBox[10, 6] = pictureBox_06_K;
            pictureBox[10, 7] = pictureBox_07_K;
            pictureBox[10, 8] = pictureBox_08_K;
            pictureBox[10, 9] = pictureBox_09_K;
            pictureBox[10, 10] = pictureBox_10_K;
            pictureBox[10, 11] = pictureBox_11_K;
            pictureBox[10, 12] = pictureBox_12_K;
            pictureBox[10, 13] = pictureBox_13_K;
            pictureBox[10, 14] = pictureBox_14_K;
            pictureBox[10, 15] = pictureBox_15_K;

            //Riga 11
            pictureBox[11, 0] = pictureBox_00_L;
            pictureBox[11, 1] = pictureBox_01_L;
            pictureBox[11, 2] = pictureBox_02_L;
            pictureBox[11, 3] = pictureBox_03_L;
            pictureBox[11, 4] = pictureBox_04_L;
            pictureBox[11, 5] = pictureBox_05_L;
            pictureBox[11, 6] = pictureBox_06_L;
            pictureBox[11, 7] = pictureBox_07_L;
            pictureBox[11, 8] = pictureBox_08_L;
            pictureBox[11, 9] = pictureBox_09_L;
            pictureBox[11, 10] = pictureBox_10_L;
            pictureBox[11, 11] = pictureBox_11_L;
            pictureBox[11, 12] = pictureBox_12_L;
            pictureBox[11, 13] = pictureBox_13_L;
            pictureBox[11, 14] = pictureBox_14_L;
            pictureBox[11, 15] = pictureBox_15_L;

            //Riga 12
            pictureBox[12, 0] = pictureBox_00_M;
            pictureBox[12, 1] = pictureBox_01_M;
            pictureBox[12, 2] = pictureBox_02_M;
            pictureBox[12, 3] = pictureBox_03_M;
            pictureBox[12, 4] = pictureBox_04_M;
            pictureBox[12, 5] = pictureBox_05_M;
            pictureBox[12, 6] = pictureBox_06_M;
            pictureBox[12, 7] = pictureBox_07_M;
            pictureBox[12, 8] = pictureBox_08_M;
            pictureBox[12, 9] = pictureBox_09_M;
            pictureBox[12, 10] = pictureBox_10_M;
            pictureBox[12, 11] = pictureBox_11_M;
            pictureBox[12, 12] = pictureBox_12_M;
            pictureBox[12, 13] = pictureBox_13_M;
            pictureBox[12, 14] = pictureBox_14_M;
            pictureBox[12, 15] = pictureBox_15_M;

            //Riga 13
            pictureBox[13, 0] = pictureBox_00_N;
            pictureBox[13, 1] = pictureBox_01_N;
            pictureBox[13, 2] = pictureBox_02_N;
            pictureBox[13, 3] = pictureBox_03_N;
            pictureBox[13, 4] = pictureBox_04_N;
            pictureBox[13, 5] = pictureBox_05_N;
            pictureBox[13, 6] = pictureBox_06_N;
            pictureBox[13, 7] = pictureBox_07_N;
            pictureBox[13, 8] = pictureBox_08_N;
            pictureBox[13, 9] = pictureBox_09_N;
            pictureBox[13, 10] = pictureBox_10_N;
            pictureBox[13, 11] = pictureBox_11_N;
            pictureBox[13, 12] = pictureBox_12_N;
            pictureBox[13, 13] = pictureBox_13_N;
            pictureBox[13, 14] = pictureBox_14_N;
            pictureBox[13, 15] = pictureBox_15_N;

            //Riga 14
            pictureBox[14, 0] = pictureBox_00_O;
            pictureBox[14, 1] = pictureBox_01_O;
            pictureBox[14, 2] = pictureBox_02_O;
            pictureBox[14, 3] = pictureBox_03_O;
            pictureBox[14, 4] = pictureBox_04_O;
            pictureBox[14, 5] = pictureBox_05_O;
            pictureBox[14, 6] = pictureBox_06_O;
            pictureBox[14, 7] = pictureBox_07_O;
            pictureBox[14, 8] = pictureBox_08_O;
            pictureBox[14, 9] = pictureBox_09_O;
            pictureBox[14, 10] = pictureBox_10_O;
            pictureBox[14, 11] = pictureBox_11_O;
            pictureBox[14, 12] = pictureBox_12_O;
            pictureBox[14, 13] = pictureBox_13_O;
            pictureBox[14, 14] = pictureBox_14_O;
            pictureBox[14, 15] = pictureBox_15_O;

            //Riga 15
            pictureBox[15, 0] = pictureBox_00_P;
            pictureBox[15, 1] = pictureBox_01_P;
            pictureBox[15, 2] = pictureBox_02_P;
            pictureBox[15, 3] = pictureBox_03_P;
            pictureBox[15, 4] = pictureBox_04_P;
            pictureBox[15, 5] = pictureBox_05_P;
            pictureBox[15, 6] = pictureBox_06_P;
            pictureBox[15, 7] = pictureBox_07_P;
            pictureBox[15, 8] = pictureBox_08_P;
            pictureBox[15, 9] = pictureBox_09_P;
            pictureBox[15, 10] = pictureBox_10_P;
            pictureBox[15, 11] = pictureBox_11_P;
            pictureBox[15, 12] = pictureBox_12_P;
            pictureBox[15, 13] = pictureBox_13_P;
            pictureBox[15, 14] = pictureBox_14_P;
            pictureBox[15, 15] = pictureBox_15_P;

            // Grafica di default
            for (int i = 0; i < 16; i++)
            {
                comboBox[i].SelectedIndex = 0;
            }

            carica_configurazione();
            carica_configurazione_3();

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);

            lang.loadLanguageTemplate(modBus_Client_.language);

            this.Title = modBus_Client_.Title;
        }


        private void carica_configurazione()
        {
            //Caricamento valori ultima sessione
            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/ComandiBit.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                SAVE_Form config = jss.Deserialize<SAVE_Form>(file_content);

                //textBoxModBusAddress.Text = config.textBoxModBusAddress_;
                //textBoxHoldingOffset.Text = config.textBoxHoldingOffset_;
                //comboBoxHoldingOffset.SelectedIndex = config.comboBoxHoldingOffset_;

                // Variabili array 
                String[] textBoxLabel_ = new String[numberOfRegisters];
                String[] textBoxVal_ = new String[numberOfRegisters];
                String[] textBoxValue_ = new String[numberOfRegisters];

                String[,] labelBitRegisters_ = new String[numberOfRegisters, 16];

                string[] comboBoxVal_ = new string[numberOfRegisters];

                textBoxLabel_ = config.textBoxLabel_;
                textBoxVal_ = config.textBoxVal_;
                textBoxValue_ = config.textBoxValue_;
                comboBoxVal_ = config.comboBoxVal_;

                checkBoxOffBit.IsChecked = config.checkBoxOffBit_;


                for (int i = 0; i < numberOfRegisters; i++)
                {
                    textBoxLabel[i].Text = textBoxLabel_[i];
                    textBoxRegister[i].Text = textBoxVal_[i];
                    textBoxValue[i].Text = textBoxValue_[i];

                    comboBox[i].SelectedIndex = comboBoxVal_[i] == "HEX" ? 1 : 0;
                }

                // Inserito qua in modo che vengano caricati i dati anche da versioni di salvataggi precedenti l'ultimo aggiornamento
                if (config.registriBloccati_)
                {
                    for (int i = 0; i < numberOfRegisters; i++)
                    {
                        textBoxLabel[i].IsEnabled = false;
                        textBoxRegister[i].IsEnabled = false;
                        textBoxValue[i].IsEnabled = false;

                        comboBox[i].IsEnabled = false;
                    }

                    textBoxModBusAddress.IsEnabled = false;
                    comboBoxHoldingOffset.IsEnabled = false;
                    textBoxHoldingOffset.IsEnabled = false;
                }

                checkBoxDisablePictureBoxCommands.IsChecked = config.disattivaComandiPictureBox_;
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
                for (int i = 0; i < numberOfRegisters; i++)
                {

                    string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/Label_ComandiBit_" + i.ToString() + ".json");

                    JavaScriptSerializer jss = new JavaScriptSerializer();
                    SAVE_Form3 config = jss.Deserialize<SAVE_Form3>(file_content);

                    String[] labelBitRegisters_ = new String[16];

                    labelBitRegisters_ = config.labelBitRegisters_;

                    for (int a = 0; a < 16; a++)
                    {
                        labelBitRegisters[i, a] = labelBitRegisters_[a];

                        // Tooltips
                        pictureBox[i, a].ToolTip = labelBitRegisters[i, a];
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        //-------------------------------------------------------------------------
        //-------------------------Funzione FORM CLOSING---------------------------
        //-------------------------------------------------------------------------

        private void Sim_Form_cs_Closing(object sender, EventArgs e)
        {
            salva_configurazione();
            modBus_Client.toolSTripMenuEnable(true);
        }

        public void salva_configurazione()
        {
            //Salvataggio valori ultima sessione
            try
            {
                SAVE_Form config = new SAVE_Form();

                config.textBoxModBusAddress_ = textBoxModBusAddress.Text;
                config.textBoxHoldingOffset_ = textBoxHoldingOffset.Text;
                config.comboBoxHoldingOffset_ = comboBoxHoldingOffset.SelectedValue.ToString().Split(' ')[1];
                config.disattivaComandiPictureBox_ = (bool)checkBoxDisablePictureBoxCommands.IsChecked;

                // Variabili array 
                String[] textBoxLabel_ = new String[numberOfRegisters];
                String[] textBoxVal_ = new String[numberOfRegisters];
                String[] textBoxValue_ = new String[numberOfRegisters];

                string[] comboBoxVal_ = new string[numberOfRegisters];

                for (int i = 0; i < numberOfRegisters; i++)
                {
                    textBoxLabel_[i] = textBoxLabel[i].Text;
                    textBoxVal_[i] = textBoxRegister[i].Text;
                    textBoxValue_[i] = textBoxValue[i].Text;

                    comboBoxVal_[i] = comboBox[i].SelectedValue.ToString().Split(' ')[1];

                }



                config.textBoxLabel_ = textBoxLabel_;
                config.textBoxVal_ = textBoxVal_;
                config.textBoxValue_ = textBoxValue_;
                config.comboBoxVal_ = comboBoxVal_;

                config.checkBoxOffBit_ = (bool)checkBoxOffBit.IsChecked;
                config.registriBloccati_ = !textBoxLabel[0].IsEnabled;

                JavaScriptSerializer jss = new JavaScriptSerializer();
                string file_content = jss.Serialize(config);

                File.WriteAllText("Json/" + pathToConfiguration + "/ComandiBit.json", file_content);
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }


        //-------------------------------------------------------------------------
        //---------------------Funzioni generali per elementi----------------------
        //-------------------------------------------------------------------------

        // COnverte un array di 16 pictureBox in un UInt16
        private UInt16 pictureToInt(int row)
        {
            UInt16 val = 0;

            for (int i = 0; i < 16; i++)
            {
                if (pictureBox[row, i].Background != Brushes.LightGray)
                {
                    val += (UInt16)((UInt16)(1) << i);
                }
            }

            return val;
        }

        //Converte un valore UInt16 in una rappresentazione tramite un array di 16 pictureBox
        private void intToPicture(int row, int value)
        {
            for (int i = 0; i < 16; i++)
            {
                if ((value & (1 << i)) > 0)
                {
                    pictureBox[row, i].Background = Brushes.Orange;
                }
                else
                {
                    pictureBox[row, i].Background = Brushes.LightGray;
                }
            }
        }

        private void picture(int row, int bit)
        {
            if (!(bool)checkBoxDisablePictureBoxCommands.IsChecked)
            {
                pictureBoxBusy.Background = Brushes.Yellow;

                int repeat = new int();

                if ((bool)checkBoxOffBit.IsChecked)
                    repeat = 2;
                else
                    repeat = 1;

                // Se ce la spunta sulla disattivazione automatica del bit ripeto due volte la seguente funzione
                for (int i = 0; i < repeat; i++)
                {
                    if (pictureBox[row, bit].Background == Brushes.LightGray)
                    {
                        pictureBox[row, bit].Background = Brushes.Orange;
                    }
                    else
                    {
                        pictureBox[row, bit].Background = Brushes.LightGray;
                        repeat = 1; // Se il bit era gia' settato da prima evito
                                    // di esguire la funzione due volte
                    }

                    UInt16 val = pictureToInt(row);
                    UInt16 address = (UInt16)(P.uint_parser(textBoxRegister[row], comboBox[row]) + P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset));

                    textBoxValue[row].Text = val.ToString();

                    ModBus.presetSingleRegister_06(byte.Parse(textBoxModBusAddress.Text), address, val);
                }

                pictureBoxBusy.Background = Brushes.LightGray;
            }
        }

        private void reset(int row)
        {
            pictureBoxBusy.Background = Brushes.Yellow;

            textBoxValue[row].Text = "0";

            for (int i = 0; i < 16; i++)
            {
                pictureBox[row, i].Background = Brushes.LightGray;
            }

            UInt16 address = (UInt16)(P.uint_parser(textBoxRegister[row], comboBox[row]) + P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset));

            ModBus.presetSingleRegister_06(byte.Parse(textBoxModBusAddress.Text), address, 0);

            pictureBoxBusy.Background = Brushes.LightGray;
        }

        private void read(int row)
        {
            pictureBoxBusy.Background = Brushes.Yellow;

            UInt16 address = (UInt16)(P.uint_parser(textBoxRegister[row], comboBox[row]) + P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset));

            UInt16 val = UInt16.Parse(ModBus.readHoldingRegister_03(byte.Parse(textBoxModBusAddress.Text), address, 1)[0]);

            intToPicture(row, val);
            textBoxValue[row].Text = val.ToString();

            pictureBoxBusy.Background = Brushes.LightGray;
        }

        private void edit(int row)
        {
            UInt16 offset = (UInt16)(P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset));
            UInt16 address = (UInt16)(P.uint_parser(textBoxRegister[row], comboBox[row]));

            try
            {
                String[] S_titleLalbel = new string[numberOfRegisters];

                for (int i = 0; i < numberOfRegisters; i++)
                {
                    S_titleLalbel[i] = textBoxLabel[i].Text;
                }

                Editor editorForm = new Editor(false, ModBus, row, this, pathToConfiguration, !textBoxLabel[0].IsEnabled, modBus_Client);
                editorForm.Show();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }



        //-------------------------------------------------------------------------
        //-----------------------Funzioni click su pictureBox----------------------
        //-------------------------------------------------------------------------

        private void pictureBox_0_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 0);
        }

        private void pictureBox_1_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 1);
        }

        private void pictureBox_2_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 2);
        }

        private void pictureBox_3_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 3);
        }

        private void pictureBox_4_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 4);
        }

        private void pictureBox_5_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 5);
        }

        private void pictureBox_6_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 6);
        }

        private void pictureBox_7_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 7);
        }

        private void pictureBox_8_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 8);
        }

        private void pictureBox_9_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 9);
        }

        private void pictureBox_10_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 10);
        }

        private void pictureBox_11_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 11);
        }

        private void pictureBox_12_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 12);
        }

        private void pictureBox_13_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 13);
        }

        private void pictureBox_14_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 14);
        }

        private void pictureBox_15_A_Click(object sender, MouseButtonEventArgs e)
        {
            picture(0, 15);
        }

        private void pictureBox_0_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 0);
        }

        private void pictureBox_1_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 1);
        }

        private void pictureBox_2_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 2);
        }

        private void pictureBox_3_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 3);
        }

        private void pictureBox_4_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 4);
        }

        private void pictureBox_5_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 5);
        }

        private void pictureBox_6_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 6);
        }

        private void pictureBox_7_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 7);
        }

        private void pictureBox_8_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 8);
        }

        private void pictureBox_9_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 9);
        }

        private void pictureBox_10_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 10);
        }

        private void pictureBox_11_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 11);
        }

        private void pictureBox_12_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 12);
        }

        private void pictureBox_13_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 13);
        }

        private void pictureBox_14_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 14);
        }

        private void pictureBox_15_B_Click(object sender, MouseButtonEventArgs e)
        {
            picture(1, 15);
        }

        private void pictureBox_0_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 0);
        }

        private void pictureBox_1_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 1);
        }

        private void pictureBox_2_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 2);
        }

        private void pictureBox_3_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 3);
        }

        private void pictureBox_4_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 4);
        }

        private void pictureBox_5_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 5);
        }

        private void pictureBox_6_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 6);
        }

        private void pictureBox_7_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 7);
        }

        private void pictureBox_8_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 8);
        }

        private void pictureBox_9_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 9);
        }

        private void pictureBox_10_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 10);
        }

        private void pictureBox_11_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 11);
        }

        private void pictureBox_12_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 12);
        }

        private void pictureBox_13_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 13);
        }

        private void pictureBox_14_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 14);
        }

        private void pictureBox_15_C_Click(object sender, MouseButtonEventArgs e)
        {
            picture(2, 15);
        }

        private void pictureBox_0_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 0);
        }

        private void pictureBox_1_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 1);
        }

        private void pictureBox_2_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 2);
        }

        private void pictureBox_3_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 3);
        }

        private void pictureBox_4_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 4);
        }

        private void pictureBox_5_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 5);
        }

        private void pictureBox_6_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 6);
        }

        private void pictureBox_7_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 7);
        }

        private void pictureBox_8_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 8);
        }

        private void pictureBox_9_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 9);
        }

        private void pictureBox_10_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 10);
        }

        private void pictureBox_11_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 11);
        }

        private void pictureBox_12_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 12);
        }

        private void pictureBox_13_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 13);
        }

        private void pictureBox_14_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 14);
        }

        private void pictureBox_15_D_Click(object sender, MouseButtonEventArgs e)
        {
            picture(3, 15);
        }

        private void pictureBox_0_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 0);
        }

        private void pictureBox_1_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 1);
        }

        private void pictureBox_2_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 2);
        }

        private void pictureBox_3_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 3);
        }

        private void pictureBox_4_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 4);
        }

        private void pictureBox_5_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 5);
        }

        private void pictureBox_6_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 6);
        }

        private void pictureBox_7_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 7);
        }

        private void pictureBox_8_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 8);
        }

        private void pictureBox_9_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 9);
        }

        private void pictureBox_10_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 10);
        }

        private void pictureBox_11_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 11);
        }

        private void pictureBox_12_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 12);
        }

        private void pictureBox_13_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 13);
        }

        private void pictureBox_14_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 14);
        }

        private void pictureBox_15_E_Click(object sender, MouseButtonEventArgs e)
        {
            picture(4, 15);
        }

        private void pictureBox_0_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 0);
        }

        private void pictureBox_1_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 1);
        }

        private void pictureBox_2_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 2);
        }

        private void pictureBox_3_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 3);
        }

        private void pictureBox_4_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 4);
        }

        private void pictureBox_5_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 5);
        }

        private void pictureBox_6_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 6);
        }

        private void pictureBox_7_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 7);
        }

        private void pictureBox_8_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 8);
        }

        private void pictureBox_9_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 9);
        }

        private void pictureBox_10_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 10);
        }

        private void pictureBox_11_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 11);
        }

        private void pictureBox_12_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 12);
        }

        private void pictureBox_13_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 13);
        }

        private void pictureBox_14_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 14);
        }

        private void pictureBox_15_F_Click(object sender, MouseButtonEventArgs e)
        {
            picture(5, 15);
        }

        private void pictureBox_0_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 0);
        }

        private void pictureBox_1_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 1);
        }

        private void pictureBox_2_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 2);
        }

        private void pictureBox_3_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 3);
        }

        private void pictureBox_4_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 4);
        }

        private void pictureBox_5_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 5);
        }

        private void pictureBox_6_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 6);
        }

        private void pictureBox_7_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 7);
        }

        private void pictureBox_8_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 8);
        }

        private void pictureBox_9_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 9);
        }

        private void pictureBox_10_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 10);
        }

        private void pictureBox_11_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 11);
        }

        private void pictureBox_12_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 12);
        }

        private void pictureBox_13_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 13);
        }

        private void pictureBox_14_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 14);
        }

        private void pictureBox_15_G_Click(object sender, MouseButtonEventArgs e)
        {
            picture(6, 15);
        }

        private void pictureBox_0_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 0);
        }

        private void pictureBox_1_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 1);
        }

        private void pictureBox_2_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 2);
        }

        private void pictureBox_3_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 3);
        }

        private void pictureBox_4_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 4);
        }

        private void pictureBox_5_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 5);
        }

        private void pictureBox_6_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 6);
        }

        private void pictureBox_7_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 7);
        }

        private void pictureBox_9_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 9);
        }

        private void pictureBox_8_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 8);
        }

        private void pictureBox_10_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 10);
        }

        private void pictureBox_11_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 11);
        }

        private void pictureBox_12_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 12);
        }

        private void pictureBox_13_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 13);
        }

        private void pictureBox_14_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 14);
        }

        private void pictureBox_15_H_Click(object sender, MouseButtonEventArgs e)
        {
            picture(7, 15);
        }

        private void buttonRead_A_Click(object sender, RoutedEventArgs e)
        {
            read(0);
        }

        private void buttonRead_B_Click(object sender, RoutedEventArgs e)
        {
            read(1);
        }

        private void buttonRead_C_Click(object sender, RoutedEventArgs e)
        {
            read(2);
        }

        private void buttonRead_D_Click(object sender, RoutedEventArgs e)
        {
            read(3);
        }

        private void buttonRead_E_Click(object sender, RoutedEventArgs e)
        {
            read(4);
        }

        private void buttonRead_F_Click(object sender, RoutedEventArgs e)
        {
            read(5);
        }

        private void buttonRead_G_Click(object sender, RoutedEventArgs e)
        {
            read(6);
        }

        private void buttonRead_H_Click(object sender, RoutedEventArgs e)
        {
            read(7);
        }

        private void buttonReset_A_Click(object sender, RoutedEventArgs e)
        {
            reset(0);
        }

        private void buttonReset_B_Click(object sender, RoutedEventArgs e)
        {
            reset(1);
        }

        private void buttonReset_C_Click(object sender, RoutedEventArgs e)
        {
            reset(2);
        }

        private void buttonReset_D_Click(object sender, RoutedEventArgs e)
        {
            reset(3);
        }

        private void buttonReset_E_Click(object sender, RoutedEventArgs e)
        {
            reset(4);
        }

        private void buttonReset_F_Click(object sender, RoutedEventArgs e)
        {
            reset(5);
        }

        private void buttonReset_G_Click(object sender, RoutedEventArgs e)
        {
            reset(6);
        }

        private void buttonReset_H_Click(object sender, RoutedEventArgs e)
        {
            reset(7);
        }

        // Classe per caricare dati dal file di configurazione json
        public class SAVE_Form
        {
            public String textBoxModBusAddress_ { get; set; }
            public String textBoxHoldingOffset_ { get; set; }
            public string comboBoxHoldingOffset_ { get; set; }

            // Variabili array 
            public String[] textBoxLabel_ { get; set; }
            public String[] textBoxVal_ { get; set; }
            public String[] textBoxValue_ { get; set; }

            public string[] comboBoxVal_ { get; set; }

            public bool checkBoxOffBit_ { get; set; }
            public bool registriBloccati_ { get; set; }
            public bool disattivaComandiPictureBox_ { get; set; }
        }

        // Classe per caricare dati dal file di configurazione json
        public class SAVE_Form3
        {
            public String[] labelBitRegisters_ { get; set; }    // ToolTips bit
        }

        private void pictureBox_0_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 0);
        }

        private void pictureBox_0_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 0);
        }

        private void pictureBox_0_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 0);
        }

        private void pictureBox_0_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 0);
        }

        private void pictureBox_1_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 1);
        }

        private void pictureBox_1_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 1);
        }

        private void pictureBox_1_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 1);
        }

        private void pictureBox_1_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 1);
        }

        private void pictureBox_2_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 2);
        }

        private void pictureBox_2_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 2);
        }

        private void pictureBox_2_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 2);
        }

        private void pictureBox_2_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 2);
        }

        private void pictureBox_3_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 3);
        }

        private void pictureBox_3_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 3);
        }

        private void pictureBox_3_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 3);
        }

        private void pictureBox_3_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 3);
        }

        private void pictureBox_4_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 4);
        }

        private void pictureBox_4_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 4);
        }

        private void pictureBox_4_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 4);
        }

        private void pictureBox_4_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 4);
        }

        private void pictureBox_5_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 5);
        }

        private void pictureBox_5_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 5);
        }

        private void pictureBox_5_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 5);
        }

        private void pictureBox_5_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 5);
        }

        private void pictureBox_6_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 6);
        }

        private void pictureBox_6_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 6);
        }

        private void pictureBox_6_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 6);
        }

        private void pictureBox_6_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 6);
        }

        private void pictureBox_7_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 7);
        }

        private void pictureBox_7_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 7);
        }

        private void pictureBox_7_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 7);
        }

        private void pictureBox_7_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 7);
        }

        private void pictureBox_8_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 8);
        }

        private void pictureBox_8_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 8);
        }

        private void pictureBox_8_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 8);
        }

        private void pictureBox_8_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 8);
        }

        private void pictureBox_9_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 9);
        }

        private void pictureBox_9_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 9);
        }

        private void pictureBox_9_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 9);
        }

        private void pictureBox_9_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 9);
        }

        private void pictureBox_10_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 10);
        }

        private void pictureBox_10_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 10);
        }

        private void pictureBox_10_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 10);
        }

        private void pictureBox_10_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 10);
        }

        private void pictureBox_11_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 11);
        }

        private void pictureBox_11_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 11);
        }

        private void pictureBox_11_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 11);
        }

        private void pictureBox_11_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 11);
        }

        private void pictureBox_12_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 12);
        }

        private void pictureBox_12_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 12);
        }

        private void pictureBox_12_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 12);
        }

        private void pictureBox_12_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 12);
        }

        private void pictureBox_13_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 13);
        }

        private void pictureBox_13_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 13);
        }

        private void pictureBox_13_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 13);
        }

        private void pictureBox_13_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 13);
        }

        private void pictureBox_14_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 14);
        }

        private void pictureBox_14_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 14);
        }

        private void pictureBox_14_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 14);
        }

        private void pictureBox_14_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 14);
        }

        private void pictureBox_15_I_Click(object sender, RoutedEventArgs e)
        {
            picture(8, 15);
        }

        private void pictureBox_15_J_Click(object sender, RoutedEventArgs e)
        {
            picture(9, 15);
        }

        private void pictureBox_15_K_Click(object sender, RoutedEventArgs e)
        {
            picture(10, 15);
        }

        private void pictureBox_15_L_Click(object sender, RoutedEventArgs e)
        {
            picture(11, 15);
        }

        private void buttonRead_I_Click(object sender, RoutedEventArgs e)
        {
            read(8);
        }

        private void buttonRead_J_Click(object sender, RoutedEventArgs e)
        {
            read(9);
        }

        private void buttonRead_K_Click(object sender, RoutedEventArgs e)
        {
            read(10);
        }

        private void buttonRead_L_Click(object sender, RoutedEventArgs e)
        {
            read(11);
        }

        private void buttonReset_I_Click(object sender, RoutedEventArgs e)
        {
            reset(8);
        }

        private void buttonReset_J_Click(object sender, RoutedEventArgs e)
        {
            reset(9);
        }

        private void buttonReset_K_Click(object sender, RoutedEventArgs e)
        {
            reset(10);
        }

        private void buttonReset_L_Click(object sender, RoutedEventArgs e)
        {
            reset(11);
        }

        private void buttonEdit_A_Click(object sender, RoutedEventArgs e)
        {
            edit(0);
        }

        private void buttonEdit_B_Click(object sender, RoutedEventArgs e)
        {
            edit(1);
        }

        private void buttonEdit_C_Click(object sender, RoutedEventArgs e)
        {
            edit(2);
        }

        private void buttonEdit_D_Click(object sender, RoutedEventArgs e)
        {
            edit(3);
        }

        private void buttonEdit_E_Click(object sender, RoutedEventArgs e)
        {
            edit(4);
        }

        private void buttonEdit_F_Click(object sender, RoutedEventArgs e)
        {
            edit(5);
        }

        private void buttonEdit_G_Click(object sender, RoutedEventArgs e)
        {
            edit(6);
        }

        private void buttonEdit_H_Click(object sender, RoutedEventArgs e)
        {
            edit(7);
        }

        private void buttonEdit_I_Click(object sender, RoutedEventArgs e)
        {
            edit(8);
        }

        private void buttonEdit_J_Click(object sender, RoutedEventArgs e)
        {
            edit(9);
        }

        private void buttonEdit_K_Click(object sender, RoutedEventArgs e)
        {
            edit(10);
        }

        private void buttonEdit_L_Click(object sender, RoutedEventArgs e)
        {
            edit(11);
        }

        private void toolStripMenuItem1_Click(object sender, RoutedEventArgs e)
        {
            carica_configurazione_3();
        }

        private void buttonReadAll_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < numberOfRegisters; i++)
            {
                read(i);
            }
        }

        private void buttonResetAll_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult dialogResult = MessageBox.Show("Reset all?", "Alert", MessageBoxButton.YesNo);

            if (dialogResult == MessageBoxResult.Yes)
            {
                for (int i = 0; i < numberOfRegisters; i++)
                {
                    reset(i);
                }
            }
            else if (dialogResult == MessageBoxResult.No)
            {
                //Code choice "No"
            }
        }

        //-----------------------------------------------------------------------------------------
        //------------------FUNZIONE PER AGGIORNARE ELEMENTI GRAFICA DA ALTRI FORM-----------------
        //-----------------------------------------------------------------------------------------

        public void setPictureBoxFromOtherThread(int row, int bit, bool value, bool busy)
        {
            if (value)
                pictureBox[row, bit].Background = Brushes.Orange;
            else
                pictureBox[row, bit].Background = Brushes.LightGray;

            if (busy)
                pictureBoxBusy.Background = Brushes.Yellow;
            else
                pictureBoxBusy.Background = Brushes.LightGray;

            textBoxValue[row].Text = pictureToInt(row).ToString();

            // Aggiorno grafica
            // Application.DoEvents();
        }

        public void setPictureBoxBusy(bool busy)
        {
            if (busy)
                pictureBoxBusy.Background = Brushes.Yellow;
            else
                pictureBoxBusy.Background = Brushes.LightGray;
        }

        public void setPictureBoxBit(int row, int bit, bool value)
        {
            if (value)
                pictureBox[row, bit].Background = Brushes.Orange;
            else
                pictureBox[row, bit].Background = Brushes.LightGray;

            textBoxValue[row].Text = pictureToInt(row).ToString();
        }

        private void pictureBox_15_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 15);
        }

        private void pictureBox_15_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 15);
        }

        private void pictureBox_15_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 15);
        }

        private void pictureBox_15_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 15);
        }

        private void pictureBox_14_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 14);
        }

        private void pictureBox_14_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 14);
        }

        private void pictureBox_14_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 14);
        }

        private void pictureBox_14_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 14);
        }

        private void pictureBox_13_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 13);
        }

        private void pictureBox_13_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 13);
        }

        private void pictureBox_13_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 13);
        }

        private void pictureBox_13_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 13);
        }

        private void pictureBox_12_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 12);
        }

        private void pictureBox_12_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 12);
        }

        private void pictureBox_12_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 12);
        }

        private void pictureBox_12_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 12);
        }

        private void pictureBox_11_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 12);
        }

        private void pictureBox_11_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 12);
        }

        private void pictureBox_11_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 12);
        }

        private void pictureBox_11_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 12);
        }

        private void pictureBox_10_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 10);
        }

        private void pictureBox_10_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 10);
        }

        private void pictureBox_10_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 10);
        }

        private void pictureBox_10_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 10);
        }

        private void pictureBox_9_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 9);
        }

        private void pictureBox_9_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 9);
        }

        private void pictureBox_9_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 9);
        }

        private void pictureBox_9_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 9);
        }

        private void pictureBox_8_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 8);
        }

        private void pictureBox_8_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 8);
        }

        private void pictureBox_8_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 8);
        }

        private void pictureBox_8_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 8);
        }

        private void pictureBox_7_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 7);
        }

        private void pictureBox_7_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 7);
        }

        private void pictureBox_7_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 7);
        }

        private void pictureBox_7_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 7);
        }

        private void pictureBox_6_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 6);
        }

        private void pictureBox_6_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 6);
        }

        private void pictureBox_6_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 6);
        }

        private void pictureBox_6_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 6);
        }

        private void pictureBox_5_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 5);
        }

        private void pictureBox_5_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 5);
        }

        private void pictureBox_5_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 5);
        }

        private void pictureBox_5_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 5);
        }

        private void pictureBox_4_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 4);
        }

        private void pictureBox_4_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 4);
        }

        private void pictureBox_4_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 4);
        }

        private void pictureBox_4_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 4);
        }

        private void pictureBox_3_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 3);
        }

        private void pictureBox_3_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 3);
        }

        private void pictureBox_3_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 3);
        }

        private void pictureBox_3_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 3);
        }

        private void pictureBox_2_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 2);
        }

        private void pictureBox_2_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 2);
        }

        private void pictureBox_2_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 2);
        }

        private void pictureBox_2_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 2);
        }

        private void pictureBox_1_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 1);
        }

        private void pictureBox_1_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 1);
        }

        private void pictureBox_1_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 1);
        }

        private void pictureBox_1_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 1);
        }

        private void pictureBox_0_M_Click(object sender, RoutedEventArgs e)
        {
            picture(12, 0);
        }

        private void pictureBox_0_N_Click(object sender, RoutedEventArgs e)
        {
            picture(13, 0);
        }

        private void pictureBox_0_O_Click(object sender, RoutedEventArgs e)
        {
            picture(14, 0);
        }

        private void pictureBox_0_P_Click(object sender, RoutedEventArgs e)
        {
            picture(15, 0);
        }

        private void buttonEdit_M_Click(object sender, RoutedEventArgs e)
        {
            edit(12);
        }

        private void buttonEdit_N_Click(object sender, RoutedEventArgs e)
        {
            edit(13);
        }

        private void buttonEdit_O_Click(object sender, RoutedEventArgs e)
        {
            edit(14);
        }

        private void buttonEdit_P_Click(object sender, RoutedEventArgs e)
        {
            edit(15);
        }

        private void buttonRead_M_Click(object sender, RoutedEventArgs e)
        {
            read(12);
        }

        private void buttonRead_N_Click(object sender, RoutedEventArgs e)
        {
            read(13);
        }

        private void buttonRead_O_Click(object sender, RoutedEventArgs e)
        {
            read(14);
        }

        private void buttonRead_P_Click(object sender, RoutedEventArgs e)
        {
            read(15);
        }

        private void buttonReset_M_Click(object sender, RoutedEventArgs e)
        {
            reset(12);
        }

        private void buttonReset_N_Click(object sender, RoutedEventArgs e)
        {
            reset(13);
        }

        private void buttonReset_O_Click(object sender, RoutedEventArgs e)
        {
            reset(14);
        }

        private void buttonReset_P_Click(object sender, RoutedEventArgs e)
        {
            reset(15);
        }

        private void bloccasbloccaTextBoxToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                textBoxLabel[i].IsEnabled = !textBoxLabel[i].IsEnabled;
                textBoxRegister[i].IsEnabled = textBoxLabel[i].IsEnabled;
                textBoxValue[i].IsEnabled = textBoxLabel[i].IsEnabled;

                comboBox[i].IsEnabled = textBoxLabel[i].IsEnabled;
            }

            textBoxModBusAddress.IsEnabled = textBoxLabel[0].IsEnabled;
            comboBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
            textBoxHoldingOffset.IsEnabled = textBoxLabel[0].IsEnabled;
        }

        private void pictureBox_15_A_MouseUp(object sender, MouseButtonEventArgs e)
        {
            MessageBox.Show("Ciao", "Info");
        }

        private void MenuItemSalvaConfigurazione_Checked(object sender, RoutedEventArgs e)
        {
            salva_configurazione();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
