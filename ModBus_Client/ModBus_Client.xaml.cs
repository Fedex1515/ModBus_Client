

// -------------------------------------------------------------------------------------------

// Copyright (c) 2021 Federico Turco

// Permission is hereby granted, free of charge, to any person
// obtaining a copy of this software and associated documentation
// files (the "Software"), to deal in the Software without
// restriction, including without limitation the rights to use,
// copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following
// conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
// OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
// HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
// OTHER DEALINGS IN THE SOFTWARE.

// -------------------------------------------------------------------------------------------

// NB: I file in pdf accessibili dal menu info sono di proprietà dei rispettivi autori

// -------------------------------------------------------------------------------------------



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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Reflection;

//Process
using System.Diagnostics;

//Sockets
using System.Net.Sockets;

//Threading per server ModBus TCP
using System.Threading;

using System.Collections;

//Porta seriale
using System.IO.Ports;

//Comandi apri/chiudi console
using System.Runtime.InteropServices;

//Libreria JSON
//using Newtonsoft.Json;
using System.IO;

using ModBusMaster_Chicco;

// Classe con funzioni di conversione DEC-HEX
using Raccolta_funzioni_parser;

// Ping
using System.Net.NetworkInformation;

// Json LIBs
using System.Web.Script.Serialization;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //NOTE Icona
        // Icona -> va messa nel Form1.Designer dopo averla caricata in Proprieta' -> Risorse
        // this.Icon = MQTT_Almaviva.Properties.Resources.IconaAlgorab;

        //-----------------------------------------------------
        //-----------------Variabili globali-------------------
        //-----------------------------------------------------

        String version = "beta";    // Eventuale etichetta, major.minor lo recupera dall'assembly
        String title = "ModBus C#";

        String defaultPathToConfiguration = "Generico";
        public String pathToConfiguration;

        SolidColorBrush colorDefaultReadCell = Brushes.LightSkyBlue;
        SolidColorBrush colorDefaultWriteCell = Brushes.LightGreen;
        SolidColorBrush colorErrorCell = Brushes.Orange;

        ObservableCollection<ModBus_Item> list_coilsTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_inputsTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_inputRegistersTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_holdingRegistersTable = new ObservableCollection<ModBus_Item>();

        System.Windows.Forms.ColorDialog colorDialogBox = new System.Windows.Forms.ColorDialog();

        Parser P = new Parser();

        // Elementi per visualizzare/nascondere la finestra della console
        bool statoConsole = true;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        //Coda per i comandi seriali da inviare
        //Queue BufferSerialeOut = new Queue();

        ModBus_Chicco ModBus;

        SerialPort serialPort = new SerialPort();

        SaveFileDialog saveFileDialogBox;
        //OpenFileDialog openFileDialogBox;

        int template_coilsOffset = 0;
        int template_inputsOffset = 0;
        int template_inputRegistersOffset = 0;
        int template_HoldingOffset = 0;

        // Le liste seguenti contengono il registro già convertito in DEC duramte il caricamento del file Template.json
        ModBus_Item[] list_template_coilsTable;
        ModBus_Item[] list_template_inputsTable;
        ModBus_Item[] list_template_inputRegistersTable;
        ModBus_Item[] list_template_holdingRegistersTable;

        // Stati loop interrogazioni
        public int pauseLoop = 1000;
        public bool loopCoils01 = false;
        public bool loopCoilsRange = false;
        public bool loopInput02 = false;
        public bool loopInputRange = false;
        public bool loopInputRegister04 = false;
        public bool loopInputRegisterRange = false;
        public bool loopHolding03 = false;
        public bool loopHoldingRange = false;
        public bool loopThreadRunning = false;

        Thread threadLoopQuery;

        public MainWindow()
        {
            InitializeComponent();

            // Creo evento di chiusura del form
            this.Closing += Form1_FormClosing;

            //this.textBoxHoldingRegisterNumber.KeyUp = new new KeyEventHandler(buttonReadHolding03_Click);

            pathToConfiguration = defaultPathToConfiguration;

            // Aspetti grafici
            comboBoxDiagnosticFunction.Items.Add("00 Return Query Data");
            comboBoxDiagnosticFunction.Items.Add("01 Restart Comunications Option");
            comboBoxDiagnosticFunction.Items.Add("02 Return Diagnostic Register");
            comboBoxDiagnosticFunction.Items.Add("03 Change ASCII Input Delimeter");
            comboBoxDiagnosticFunction.Items.Add("04 Force Listen Only Mode");
            comboBoxDiagnosticFunction.Items.Add("10 Clear Counters and Diagnostic Register");
            comboBoxDiagnosticFunction.Items.Add("11 Return Bus Message Count");
            comboBoxDiagnosticFunction.Items.Add("12 Return Bus Comunication Error Count");
            comboBoxDiagnosticFunction.Items.Add("13 Return Bus Exception Error Count");
            comboBoxDiagnosticFunction.Items.Add("14 Return Slave Message Count");
            comboBoxDiagnosticFunction.Items.Add("15 Return Slave No Response Count");
            comboBoxDiagnosticFunction.Items.Add("16 Return Slave NAK Count");
            comboBoxDiagnosticFunction.Items.Add("17 Return Slave Busy Count");
            comboBoxDiagnosticFunction.Items.Add("20 Clear Overrun Counter and Flag");

            pictureBoxSerial.Background = Brushes.LightGray;
            pictureBoxTcp.Background = Brushes.LightGray;

            dataGridViewCoils.ItemsSource = list_coilsTable;
            dataGridViewInput.ItemsSource = list_inputsTable;
            dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;
            dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

            // Aspetti grafici di default
            comboBoxSerialSpeed.SelectedIndex = 7;
            comboBoxSerialParity.SelectedIndex = 0;
            comboBoxSerialStop.SelectedIndex = 0;

            textBoxTcpClientIpAddress.Text = "192.168.1.100";
            textBoxTcpClientPort.Text = "502";

            comboBoxCoilsRegistri.SelectedIndex = 0;
            comboBoxCoilsOffset.SelectedIndex = 0;
            textBoxCoilsOffset.Text = "0";

            comboBoxInputRegistri.SelectedIndex = 0;
            comboBoxInputOffset.SelectedIndex = 0;
            textBoxInputOffset.Text = "0";

            comboBoxInputRegOffset.SelectedIndex = 0;
            comboBoxInputRegValori.SelectedIndex = 0;
            comboBoxInputRegRegistri.SelectedIndex = 0;
            textBoxInputRegOffset.Text = "0";

            comboBoxHoldingRegistri.SelectedIndex = 0;
            comboBoxHoldingValori.SelectedIndex = 0;
            comboBoxHoldingOffset.SelectedIndex = 0;
            textBoxHoldingOffset.Text = "0";

            comboBoxCoilsAddress01.SelectedIndex = 0;
            comboBoxCoilsRange_A.SelectedIndex = 0;
            comboBoxCoilsRange_B.SelectedIndex = 0;
            comboBoxCoilsAddress05.SelectedIndex = 0;
            comboBoxCoilsAddress05_b.SelectedIndex = 0;
            //comboBoxCoilsValue05.SelectedIndex = 0;
            //comboBoxCoilsValue05_b.SelectedIndex = 0;
            comboBoxCoilsAddress15_A.SelectedIndex = 0;
            comboBoxCoilsAddress15_B.SelectedIndex = 0;
            //comboBoxCoilsValue15.SelectedIndex = 0;

            comboBoxInputAddress02.SelectedIndex = 0;
            comboBoxInputRange_A.SelectedIndex = 0;
            comboBoxInputRange_B.SelectedIndex = 0;

            comboBoxInputRegisterAddress04.SelectedIndex = 0;
            comboBoxInputRegisterRange_A.SelectedIndex = 0;
            comboBoxInputRegisterRange_B.SelectedIndex = 0;

            comboBoxHoldingAddress03.SelectedIndex = 0;
            comboBoxHoldingRange_A.SelectedIndex = 0;
            comboBoxHoldingRange_B.SelectedIndex = 0;
            comboBoxHoldingAddress06.SelectedIndex = 0;
            comboBoxHoldingValue06.SelectedIndex = 0;
            comboBoxHoldingAddress06_b.SelectedIndex = 0;
            comboBoxHoldingValue06_b.SelectedIndex = 0;
            comboBoxHoldingAddress16_A.SelectedIndex = 0;
            comboBoxHoldingAddress16_B.SelectedIndex = 0;
            comboBoxHoldingValue16.SelectedIndex = 0;


            pictureBoxRunningAs.Background = Brushes.LightGray;
            pictureBoxIsSending.Background = Brushes.LightGray;
            pictureBoxIsResponding.Background = Brushes.LightGray;

            richTextBoxStatus.AppendText("\n");

            radioButtonModeSerial.IsChecked = true;

            checkBoxAddLinesToEnd.Visibility = Visibility.Hidden;
        }

        private void Form1_FormClosing(object sender, EventArgs e)
        {
            salva_configurazione(false);

            try
            {
                if (ModBus != null)
                {
                    ModBus.close(); // Se non attivo niente ModBUs risulta null
                }
            }
            catch
            {

            }

            try
            {
                if (threadLoopQuery != null)
                {
                    if (threadLoopQuery.IsAlive)
                    {
                        threadLoopQuery.Abort();
                    }
                }
            }
            catch
            {

            }
        }

        private void Form1_Load(object sender, RoutedEventArgs e)
        {
            Console.WriteLine(this.Title + "\n");

            var version_ = Assembly.GetEntryAssembly().GetName().Version;

            this.Title = "ModBus Client " + version_.ToString().Split('.')[0] + "." + version_.ToString().Split('.')[1] + " " + version;

            try
            {
                // Aggiornamento lista porte seriale
                string[] SerialPortList = System.IO.Ports.SerialPort.GetPortNames();
                // comboBoxSerialPort.Items.Add("Seleziona porta seriale ...");

                foreach (String port in SerialPortList)
                {
                    comboBoxSerialPort.Items.Add(port);
                }

                comboBoxSerialPort.SelectedIndex = 0;
            }
            catch
            {
                Console.WriteLine("Nessuna porta seriale trovata");
            }

            carica_configurazione();

            Thread updateTables = new Thread(new ThreadStart(genera_tabelle_registri));
            updateTables.IsBackground = true;
            updateTables.Start();

            if (!(bool)(checkBoxCreateTableAtBoot.IsChecked))
            {
                // Nascondo gli elementi goto index

                // Coils
                labelGoToCoils.Visibility = Visibility.Hidden;
                textBoxGoToCoilAddress.Visibility = Visibility.Hidden;
                buttonGoToCoilAddress.Visibility = Visibility.Hidden;
                stackGoToCoils.Visibility = Visibility.Hidden;

                // Inputs
                labelGoToInput.Visibility = Visibility.Hidden;
                textBoxGoToInputAddress.Visibility = Visibility.Hidden;
                buttonGoToInputAddress.Visibility = Visibility.Hidden;
                stackGoToInput.Visibility = Visibility.Hidden;

                // Input Register
                labelGoToInputRegister.Visibility = Visibility.Hidden;
                textBoxGoToInputRegisterAddress.Visibility = Visibility.Hidden;
                buttonGoToInputRegisterAddress.Visibility = Visibility.Hidden;
                stackGoToInputRegister.Visibility = Visibility.Hidden;

                // Holding registers
                labelGoToHolding.Visibility = Visibility.Hidden;
                textBoxGoToHoldingAddress.Visibility = Visibility.Hidden;
                buttonGoToHoldingAddress.Visibility = Visibility.Hidden;
                stackGoToHolding.Visibility = Visibility.Hidden;
            }


            // ToolTips
            // toolTipHelp.SetToolTip(this.comboBoxHoldingAddress03, "Formato indirizzo\nDEC: decimale\nHEX: esadecimale (inserire il valore senza 0x)");
        }

        private void radioButtonModeSerial_CheckedChanged(object sender, RoutedEventArgs e)
        {

            // Serial ON

            // radioButtonModeASCII.IsEnabled = radioButtonModeSerial.IsChecked;
            // radioButtonModeRTU.IsEnabled = radioButtonModeSerial.IsChecked;

            comboBoxSerialParity.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialPort.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialSpeed.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialStop.IsEnabled = (bool)radioButtonModeSerial.IsChecked;

            buttonUpdateSerialList.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            buttonSerialActive.IsEnabled = (bool)radioButtonModeSerial.IsChecked;


            // Tcp OFF
            // radioButtonTcpSlave.IsEnabled = !radioButtonModeSerial.IsChecked;

            richTextBoxStatus.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            buttonTcpActive.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;

            if ((bool)radioButtonModeSerial.IsChecked)
            {
                textBoxTcpClientIpAddress.IsEnabled = false;
                textBoxTcpClientPort.IsEnabled = false;
            }
            else
            {

                textBoxTcpClientIpAddress.IsEnabled = true;
                textBoxTcpClientPort.IsEnabled = true;

            }
        }



        private void buttonSerialActive_Click(object sender, RoutedEventArgs e)
        {
            if (pictureBoxSerial.Background == Brushes.LightGray)
            {
                // Attivazione comunicazione seriale
                pictureBoxSerial.Background = Brushes.Lime;
                pictureBoxRunningAs.Background = Brushes.Lime;


                buttonSerialActive.Content = "Disconnect";
                // holdingSuiteToolStripMenuItem.IsEnabled = true;

                menuItemToolBit.IsEnabled = true;
                menuItemToolWord.IsEnabled = true;
                menuItemToolByte.IsEnabled = true;
                salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = false;
                caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = false;
                gestisciDatabaseToolStripMenuItem.IsEnabled = false;
                 

                try
                {
                    // ---------------------------------------------------------------------------------
                    // ----------------------Apertura comunicazione seriale-----------------------------
                    // ---------------------------------------------------------------------------------

                    // Create a new SerialPort object with default settings.
                    serialPort = new SerialPort();

                    serialPort.PortName = comboBoxSerialPort.SelectedItem.ToString();

                    // debug
                    //Console.WriteLine(comboBoxSerialSpeed.SelectedValue.ToString());
                    //Console.WriteLine(comboBoxSerialSpeed.SelectedItem.ToString());

                    serialPort.BaudRate = int.Parse(comboBoxSerialSpeed.SelectedValue.ToString().Split(' ')[1]);

                    // DEBUG
                    //Console.WriteLine("comboBoxSerialParity.SelectedIndex:" + comboBoxSerialParity.SelectedIndex.ToString());

                    switch (comboBoxSerialParity.SelectedIndex)
                    {
                        case 0:
                            serialPort.Parity = Parity.None;
                            break;
                        case 1:
                            serialPort.Parity = Parity.Even;
                            break;
                        case 2:
                            serialPort.Parity = Parity.Odd;
                            break;
                        default:
                            serialPort.Parity = Parity.None;
                            break;
                    }

                    serialPort.DataBits = 8;

                    // DEBUG
                    Console.WriteLine("comboBoxSerialStop.SelectedIndex:" + comboBoxSerialStop.SelectedIndex.ToString());

                    switch (comboBoxSerialStop.SelectedIndex)
                    {
                        case 0:
                            serialPort.StopBits = StopBits.One;
                            break;
                        case 1:
                            serialPort.StopBits = StopBits.OnePointFive;
                            break;
                        case 2:
                            serialPort.StopBits = StopBits.Two;
                            break;
                        default:
                            serialPort.StopBits = StopBits.One;
                            break;
                    }

                    serialPort.Handshake = Handshake.None;

                    // Timeout porta
                    serialPort.ReadTimeout = 50;
                    serialPort.WriteTimeout = 50;

                    ModBus = new ModBus_Chicco(serialPort, textBoxTcpClientIpAddress.Text, textBoxTcpClientPort.Text, "RTU", pictureBoxIsResponding, pictureBoxIsSending, richTextBoxOutgoingPackets, richTextBoxIncomingPackets);

                    ModBus.open();

                    serialPort.Open();
                    richTextBoxAppend(richTextBoxStatus, "Connected to " + comboBoxSerialPort.SelectedItem.ToString());

                    radioButtonModeSerial.IsEnabled = false;
                    radioButtonModeTcp.IsEnabled = false;


                    comboBoxSerialPort.IsEnabled = false;
                    comboBoxSerialSpeed.IsEnabled = false;
                    comboBoxSerialParity.IsEnabled = false;
                    comboBoxSerialStop.IsEnabled = false;
                }
                catch(Exception err)
                {
                    pictureBoxSerial.Background = Brushes.LightGray;
                    pictureBoxRunningAs.Background = Brushes.LightGray;


                    buttonSerialActive.Content = "Connect";
                    //holdingSuiteToolStripMenuItem.IsEnabled = false;

                    menuItemToolBit.IsEnabled = false;
                    menuItemToolWord.IsEnabled = false;
                    menuItemToolByte.IsEnabled = false;
                    salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = true;
                    caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = true;
                    gestisciDatabaseToolStripMenuItem.IsEnabled = true;

                    comboBoxSerialPort.IsEnabled = true;
                    comboBoxSerialSpeed.IsEnabled = true;
                    comboBoxSerialParity.IsEnabled = true;
                    comboBoxSerialStop.IsEnabled = true;

                    Console.WriteLine("Errore apertura porta seriale");
                    Console.WriteLine(err);

                    richTextBoxAppend(richTextBoxStatus, "Failed to connect");

                }
            }
            else
            {
                // Disattivazione comunicazione seriale
                pictureBoxSerial.Background = Brushes.LightGray;
                pictureBoxRunningAs.Background = Brushes.LightGray;


                buttonSerialActive.Content = "Connect";

                radioButtonModeSerial.IsEnabled = true;
                radioButtonModeTcp.IsEnabled = true;

                menuItemToolBit.IsEnabled = false;
                menuItemToolWord.IsEnabled = false;
                menuItemToolByte.IsEnabled = false;
                salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = true;
                caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = true;
                gestisciDatabaseToolStripMenuItem.IsEnabled = true;

                comboBoxSerialPort.IsEnabled = true;
                comboBoxSerialSpeed.IsEnabled = true;
                comboBoxSerialParity.IsEnabled = true;
                comboBoxSerialStop.IsEnabled = true;

                // ---------------------------------------------------------------------------------
                // ----------------------Chiusura comunicazione seriale-----------------------------
                // ---------------------------------------------------------------------------------
                serialPort.Close();
                richTextBoxAppend(richTextBoxStatus, "Port closed");
            }
        }

        private void buttonUpdateSerialList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] SerialPortList = System.IO.Ports.SerialPort.GetPortNames();

                //comboBoxSerialPort.Items.Add("Seleziona porta seriale ...");
                comboBoxSerialPort.Items.Clear();

                foreach(String port in SerialPortList)
                {
                    comboBoxSerialPort.Items.Add(port);
                }
                
                comboBoxSerialPort.SelectedIndex = 0;
            }
            catch
            {
                Console.WriteLine("Nessuna porta seriale disponibile");
            }
        }

        // Visualizza console programma da menu tendina
        private void apriConsoleToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_SHOW);

            statoConsole = true;
        }

        // Nasconde console programma da menu tendina
        private void chiudiConsoleToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            statoConsole = false;
        }

        // ----------------------------------------------------------------------------------
        // ---------------------------SALVATAGGIO CONFIGURAZIONE-----------------------------
        // ----------------------------------------------------------------------------------

        public void salva_configurazione(bool alert)   //Se alert true visualizza un messaggio di info salvataggio avvenuto
        {
            //DEBUG
            //MessageBox.Show("Salvataggio configurazione");

            try
            {
                // Caricamento variabili
                var config = new SAVE();

                config.modbusAddress = textBoxModbusAddress.Text;
                config.usingSerial = (bool)radioButtonModeSerial.IsChecked;

                //config.serialMaster = radioButtonSerialMaster.IsChecked;
                //config.serialRTU = radioButtonModeRTU.IsChecked;

                // Serial port
                config.serialPort = comboBoxSerialPort.SelectedIndex;
                config.serialSpeed = comboBoxSerialSpeed.SelectedIndex;
                config.serialParity = comboBoxSerialParity.SelectedIndex;
                config.serialStop = comboBoxSerialStop.SelectedIndex;

                // TCP
                config.tcpClientIpAddress = textBoxTcpClientIpAddress.Text;
                config.tcpClientPort = textBoxTcpClientPort.Text;
                //config.tcpServerIpAddress = textBoxTcpServerIpAddress.Text;
                //config.tcpServerPort = textBoxTcpServerPort.Text;

                // GRAFICA
                // TabPage1 (Coils)
                config.textBoxCoilsAddress01 = textBoxCoilsAddress01.Text;
                config.textBoxCoilNumber = textBoxCoilNumber.Text;
                config.textBoxCoilsRange_A = textBoxCoilsRange_A.Text;
                config.textBoxCoilsRange_B = textBoxCoilsRange_B.Text;
                config.textBoxCoilsAddress05 = textBoxCoilsAddress05.Text;
                config.textBoxCoilsValue05 = textBoxCoilsValue05.Text;
                config.textBoxCoilsAddress15_A = textBoxCoilsAddress15_A.Text;
                config.textBoxCoilsAddress15_B = textBoxCoilsAddress15_B.Text;
                config.textBoxCoilsValue15 = textBoxCoilsValue15.Text;
                config.textBoxGoToCoilAddress = textBoxGoToCoilAddress.Text;

                // TabPage2 (inputs)
                config.textBoxInputAddress02 = textBoxInputAddress02.Text;
                config.textBoxInputNumber = textBoxInputNumber.Text;
                config.textBoxInputRange_A = textBoxInputRange_A.Text;
                config.textBoxInputRange_B = textBoxInputRange_B.Text;
                config.textBoxGoToInputAddress = textBoxGoToInputAddress.Text;

                // TabPage3 (input registers)
                config.textBoxInputRegisterAddress04 = textBoxInputRegisterAddress04.Text;
                config.textBoxInputRegisterNumber = textBoxInputRegisterNumber.Text;
                config.textBoxInputRegisterRange_A = textBoxInputRegisterRange_A.Text;
                config.textBoxInputRegisterRange_B = textBoxInputRegisterRange_B.Text;
                config.textBoxGoToInputRegisterAddress = textBoxGoToInputRegisterAddress.Text;

                // TabPage4 (holding registers)
                config.textBoxHoldingAddress03 = textBoxHoldingAddress03.Text;
                config.textBoxHoldingRegisterNumber = textBoxHoldingRegisterNumber.Text;
                config.textBoxHoldingRange_A = textBoxHoldingRange_A.Text;
                config.textBoxHoldingRange_B = textBoxHoldingRange_B.Text;
                config.textBoxHoldingAddress06 = textBoxHoldingAddress06.Text;
                config.textBoxHoldingValue06 = textBoxHoldingValue06.Text;
                config.textBoxHoldingAddress16_A = textBoxHoldingAddress16_A.Text;
                config.textBoxHoldingAddress16_B = textBoxHoldingAddress16_B.Text;
                config.textBoxHoldingValue16 = textBoxHoldingValue16.Text;
                config.textBoxGoToHoldingAddress = textBoxGoToHoldingAddress.Text;

                config.statoConsole = statoConsole;

                // Funzioni aggiunte in seguito
                config.comboBoxCoilsAddress01_ = comboBoxCoilsAddress01.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsRange_A_ = comboBoxCoilsRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsRange_B_ = comboBoxCoilsRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsAddress05_ = comboBoxCoilsAddress05.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsValue05_ = "DEC"; // comboBoxCoilsValue05.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsAddress15_A_ = comboBoxCoilsAddress15_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsAddress15_B_ = comboBoxCoilsAddress15_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsValue15_ = "DEC"; // comboBoxCoilsValue15.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputAddress02_ = comboBoxInputAddress02.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRange_A_ = comboBoxInputRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRange_B_ = comboBoxInputRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegisterAddress04_ = comboBoxInputRegisterAddress04.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegisterRange_A_ = comboBoxInputRegisterRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegisterRange_B_ = comboBoxInputRegisterRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress03_ = comboBoxHoldingAddress03.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingRange_A_ = comboBoxHoldingRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingRange_B_ = comboBoxHoldingRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress06_ = comboBoxHoldingAddress06.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingValue06_ = comboBoxHoldingValue06.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress16_A_ = comboBoxHoldingAddress16_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress16_B_ = comboBoxHoldingAddress16_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingValue16_ = comboBoxHoldingValue16.SelectedValue.ToString().Split(' ')[1];

                //comboBox visualizzazione tabelle
                config.comboBoxCoilsRegistri_ = comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegistri_ = comboBoxInputRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegRegistri_ = comboBoxInputRegRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegValori_ = comboBoxInputRegValori.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingRegistri_ = comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingValori_ = comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1];

                config.comboBoxCoilsOffset_ = comboBoxCoilsOffset.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputOffset_ = comboBoxInputOffset.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegOffset_ = comboBoxInputRegOffset.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingOffset_ = comboBoxHoldingOffset.SelectedValue.ToString().Split(' ')[1];

                // Funzioni doppie per force coil e preset holding aggiunte dopo
                config.comboBoxCoilsAddress05_b_ = comboBoxCoilsAddress05_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxCoilsAddress05_b_ = textBoxCoilsAddress05_b.Text;
                config.comboBoxCoilsValue05_b_ = "DEC"; //comboBoxCoilsValue05_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxCoilsValue05_b_ = textBoxCoilsValue05_b.Text;

                config.comboBoxHoldingAddress06_b_ = comboBoxHoldingAddress06_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxHoldingAddress06_b_ = textBoxHoldingAddress06_b.Text;
                config.comboBoxHoldingValue06_b_ = comboBoxHoldingValue06_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxHoldingValue06_b_ = textBoxHoldingValue06_b.Text;

                config.textBoxCoilsOffset_ = textBoxCoilsOffset.Text;
                config.textBoxInputOffset_ = textBoxInputOffset.Text;
                config.textBoxInputRegOffset_ = textBoxInputRegOffset.Text;
                config.textBoxHoldingOffset_ = textBoxHoldingOffset.Text;

                config.checkBoxUseOffsetInTables_ = (bool)checkBoxUseOffsetInTables.IsChecked;
                config.checkBoxUseOffsetInTextBox_ = (bool)checkBoxUseOffsetInTextBox.IsChecked;
                config.checkBoxFollowModbusProtocol_ = (bool)checkBoxFollowModbusProtocol.IsChecked;
                config.checkBoxCreateTableAtBoot_ = (bool)checkBoxCreateTableAtBoot.IsChecked;
                //config.checkBoxSavePackets_ = (bool)checkBoxSavePackets.IsChecked;
                config.checkBoxCloseConsolAfterBoot_ = (bool)checkBoxCloseConsolAfterBoot.IsChecked;
                config.checkBoxCellColorMode_ = (bool)checkBoxCellColorMode.IsChecked;
                config.checkBoxViewTableWithoutOffset_ = (bool)checkBoxViewTableWithoutOffset.IsChecked;

                //config.textBoxSaveLogPath_ = textBoxSaveLogPath.Text;


                config.comboBoxDiagnosticFunction_ = comboBoxDiagnosticFunction.SelectedValue.ToString().Split(' ')[1];

                config.textBoxDiagnosticFunctionManual_ = textBoxDiagnosticFunctionManual.Text;

                config.colorDefaultReadCell_ = colorDefaultReadCell.ToString();
                config.colorDefaultWriteCell_ = colorDefaultWriteCell.ToString();
                config.colorErrorCell_ = colorErrorCell.ToString();

                config.TextBoxPollingInterval_ = TextBoxPollingInterval.Text;
                config.TextBoxPollingInterval_ = TextBoxPollingInterval.Text;

                config.CheckBoxSendValuesOnEditCoillsTable_ = (bool)CheckBoxSendValuesOnEditCoillsTable.IsChecked;
                config.CheckBoxSendValuesOnEditHoldingTable_ = (bool)CheckBoxSendValuesOnEditHoldingTable.IsChecked;

                var jss = new JavaScriptSerializer();
                jss.RecursionLimit = 1000;
                string file_content = jss.Serialize(config);

                File.WriteAllText("Json/" + pathToConfiguration + "/CONFIGURAZIONE.json", file_content);

                if (alert)
                    MessageBox.Show("Salvataggio configurazione avvenuto. Al prossimo avvio verrà caricata la configurazione corrente.", "Info");
                Console.WriteLine("Salvata configurazione");
            }
            catch(Exception err)
            {
                Console.WriteLine("Errore salvataggio configurazione");
                Console.WriteLine(err);
            }
        }

        // ----------------------------------------------------------------------------------
        // ---------------------------CARICAMENTO CONFIGURAZIONE-----------------------------
        // ----------------------------------------------------------------------------------

        public void carica_configurazione()
        {
            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/CONFIGURAZIONE.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                SAVE config = jss.Deserialize<SAVE>(file_content);

                textBoxModbusAddress.Text = config.modbusAddress;

                // Scheda configurazione seriale
                radioButtonModeSerial.IsChecked = config.usingSerial;
                radioButtonModeTcp.IsChecked = !config.usingSerial;

                //radioButtonSerialMaster.IsChecked = config.serialMaster;
                //radioButtonSerialSlave.IsChecked = !config.serialMaster;
                //radioButtonModeRTU.IsChecked = config.serialRTU;
                //radioButtonModeASCII.IsChecked = !config.serialRTU;

                comboBoxSerialSpeed.SelectedIndex = config.serialSpeed;
                comboBoxSerialParity.SelectedIndex = config.serialParity;
                comboBoxSerialStop.SelectedIndex = config.serialStop;

                // Scheda configurazione TCP
                //radioButtonTcpMaster.IsChecked = config.serialMaster;
                //radioButtonTcpSlave.IsChecked = !config.serialMaster;

                textBoxTcpClientIpAddress.Text = config.tcpClientIpAddress;
                textBoxTcpClientPort.Text = config.tcpClientPort;
                //textBoxTcpServerIpAddress.Text = config.tcpServerIpAddress;
                //textBoxTcpServerPort.Text = config.tcpServerPort;

                comboBoxSerialPort.SelectedIndex = config.serialPort;

                // GRAFICA
                // TabPage1 (Coils)
                textBoxCoilsAddress01.Text = config.textBoxCoilsAddress01;
                textBoxCoilNumber.Text = config.textBoxCoilNumber;
                textBoxCoilsRange_A.Text = config.textBoxCoilsRange_A;
                textBoxCoilsRange_B.Text = config.textBoxCoilsRange_B;
                textBoxCoilsAddress05.Text = config.textBoxCoilsAddress05;
                textBoxCoilsValue05.Text = config.textBoxCoilsValue05;
                textBoxCoilsAddress15_A.Text = config.textBoxCoilsAddress15_A;
                textBoxCoilsAddress15_B.Text = config.textBoxCoilsAddress15_B;
                textBoxCoilsValue15.Text = config.textBoxCoilsValue15;
                textBoxGoToCoilAddress.Text = config.textBoxGoToCoilAddress;


                // TabPage2 (inputs)
                textBoxInputAddress02.Text = config.textBoxInputAddress02;
                textBoxInputNumber.Text = config.textBoxInputNumber;
                textBoxInputRange_A.Text = config.textBoxInputRange_A;
                textBoxInputRange_B.Text = config.textBoxInputRange_B;
                textBoxGoToInputAddress.Text = config.textBoxGoToInputAddress;

                // TabPage3 (input registers)
                textBoxInputRegisterAddress04.Text = config.textBoxInputRegisterAddress04;
                textBoxInputRegisterNumber.Text = config.textBoxInputRegisterNumber;
                textBoxInputRegisterRange_A.Text = config.textBoxInputRegisterRange_A;
                textBoxInputRegisterRange_B.Text = config.textBoxInputRegisterRange_B;
                textBoxGoToInputRegisterAddress.Text = config.textBoxGoToInputRegisterAddress;

                // TabPage4 (holding registers)
                textBoxHoldingAddress03.Text = config.textBoxHoldingAddress03;
                textBoxHoldingRegisterNumber.Text = config.textBoxHoldingRegisterNumber;
                textBoxHoldingRange_A.Text = config.textBoxHoldingRange_A;
                textBoxHoldingRange_B.Text = config.textBoxHoldingRange_B;
                textBoxHoldingAddress06.Text = config.textBoxHoldingAddress06;
                textBoxHoldingValue06.Text = config.textBoxHoldingValue06;
                textBoxHoldingAddress16_A.Text = config.textBoxHoldingAddress16_A;
                textBoxHoldingAddress16_B.Text = config.textBoxHoldingAddress16_B;
                textBoxHoldingValue16.Text = config.textBoxHoldingValue16;
                textBoxGoToHoldingAddress.Text = config.textBoxGoToHoldingAddress;

                statoConsole = config.statoConsole;

                // Funzioni aggiunte in seguito
                comboBoxCoilsAddress01.SelectedIndex = config.comboBoxCoilsAddress01_ == "HEX" ? 1 : 0;
                comboBoxCoilsRange_A.SelectedIndex = config.comboBoxCoilsRange_A_ == "HEX" ? 1 : 0;
                comboBoxCoilsRange_B.SelectedIndex = config.comboBoxCoilsRange_B_ == "HEX" ? 1 : 0;
                comboBoxCoilsAddress05.SelectedIndex = config.comboBoxCoilsAddress05_ == "HEX" ? 1 : 0;
                //comboBoxCoilsValue05.SelectedIndex = config.comboBoxCoilsValue05_ == "HEX" ? 1 : 0;
                comboBoxCoilsAddress15_A.SelectedIndex = config.comboBoxCoilsAddress15_A_ == "HEX" ? 1 : 0;
                comboBoxCoilsAddress15_B.SelectedIndex = config.comboBoxCoilsAddress15_B_ == "HEX" ? 1 : 0;
                //comboBoxCoilsValue15.SelectedIndex = config.comboBoxCoilsValue15_ == "HEX" ? 1 : 0;
                comboBoxInputAddress02.SelectedIndex = config.comboBoxInputAddress02_ == "HEX" ? 1 : 0;
                comboBoxInputRange_A.SelectedIndex = config.comboBoxInputRange_A_ == "HEX" ? 1 : 0;
                comboBoxInputRange_B.SelectedIndex = config.comboBoxInputRange_B_ == "HEX" ? 1 : 0;
                comboBoxInputRegisterAddress04.SelectedIndex = config.comboBoxInputRegisterAddress04_ == "HEX" ? 1 : 0;
                comboBoxInputRegisterRange_A.SelectedIndex = config.comboBoxInputRegisterRange_A_ == "HEX" ? 1 : 0;
                comboBoxInputRegisterRange_B.SelectedIndex = config.comboBoxInputRegisterRange_B_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress03.SelectedIndex = config.comboBoxHoldingAddress03_ == "HEX" ? 1 : 0;
                comboBoxHoldingRange_A.SelectedIndex = config.comboBoxHoldingRange_A_ == "HEX" ? 1 : 0;
                comboBoxHoldingRange_B.SelectedIndex = config.comboBoxHoldingRange_B_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress06.SelectedIndex = config.comboBoxHoldingAddress06_ == "HEX" ? 1 : 0;
                comboBoxHoldingValue06.SelectedIndex = config.comboBoxHoldingValue06_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress16_A.SelectedIndex = config.comboBoxHoldingAddress16_A_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress16_B.SelectedIndex = config.comboBoxHoldingAddress16_B_ == "HEX" ? 1 : 0;
                comboBoxHoldingValue16.SelectedIndex = config.comboBoxHoldingValue16_ == "HEX" ? 1 : 0;

                comboBoxCoilsOffset.SelectedIndex = config.comboBoxCoilsOffset_ == "HEX" ? 1 : 0;
                comboBoxInputOffset.SelectedIndex = config.comboBoxInputOffset_ == "HEX" ? 1 : 0;
                comboBoxInputRegOffset.SelectedIndex = config.comboBoxInputRegOffset_ == "HEX" ? 1 : 0;
                comboBoxHoldingOffset.SelectedIndex = config.comboBoxHoldingOffset_ == "HEX" ? 1 : 0;

                //comboBox visualizzazione tabelle
                comboBoxCoilsRegistri.SelectedIndex = config.comboBoxCoilsRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegistri.SelectedIndex = config.comboBoxInputRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegRegistri.SelectedIndex = config.comboBoxInputRegRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegValori.SelectedIndex = config.comboBoxInputRegValori_ == "HEX" ? 1 : 0;
                comboBoxHoldingRegistri.SelectedIndex = config.comboBoxHoldingRegistri_ == "HEX" ? 1 : 0;
                comboBoxHoldingValori.SelectedIndex = config.comboBoxHoldingValori_ == "HEX" ? 1 : 0;

                textBoxCoilsOffset.Text = config.textBoxCoilsOffset_;
                textBoxInputOffset.Text = config.textBoxInputOffset_;
                textBoxInputRegOffset.Text = config.textBoxInputRegOffset_;
                textBoxHoldingOffset.Text = config.textBoxHoldingOffset_;

                // Funzioni doppie per force coil e preset holding aggiunte dopo
                comboBoxCoilsAddress05_b.SelectedIndex = config.comboBoxCoilsAddress05_b_ == "HEX" ? 1 : 0;
                textBoxCoilsAddress05_b.Text = config.textBoxCoilsAddress05_b_;
                //comboBoxCoilsValue05_b.SelectedIndex = config.comboBoxCoilsValue05_b_ == "HEX" ? 1 : 0;
                textBoxCoilsValue05_b.Text = config.textBoxCoilsValue05_b_;

                comboBoxHoldingAddress06_b.SelectedIndex = config.comboBoxHoldingAddress06_b_ == "HEX" ? 1 : 0;
                textBoxHoldingAddress06_b.Text = config.textBoxHoldingAddress06_b_;
                comboBoxHoldingValue06_b.SelectedIndex = config.comboBoxHoldingValue06_b_ == "HEX" ? 1 : 0;
                textBoxHoldingValue06_b.Text = config.textBoxHoldingValue06_b_;

                textBoxCoilsOffset.Text = config.textBoxCoilsOffset_;
                textBoxInputOffset.Text = config.textBoxInputOffset_;
                textBoxInputRegOffset.Text = config.textBoxInputRegOffset_;
                textBoxHoldingOffset.Text = config.textBoxHoldingOffset_;

                checkBoxUseOffsetInTables.IsChecked = config.checkBoxUseOffsetInTables_;
                checkBoxUseOffsetInTextBox.IsChecked = config.checkBoxUseOffsetInTextBox_;
                checkBoxFollowModbusProtocol.IsChecked = config.checkBoxFollowModbusProtocol_;
                checkBoxCreateTableAtBoot.IsChecked = config.checkBoxCreateTableAtBoot_;
                //checkBoxSavePackets.IsChecked = config.checkBoxSavePackets_;
                checkBoxCloseConsolAfterBoot.IsChecked = config.checkBoxCloseConsolAfterBoot_;
                checkBoxCellColorMode.IsChecked = config.checkBoxCellColorMode_;
                checkBoxViewTableWithoutOffset.IsChecked = config.checkBoxViewTableWithoutOffset_;

                //textBoxSaveLogPath.Text = config.textBoxSaveLogPath_;

                comboBoxDiagnosticFunction.SelectedIndex = 0;

                textBoxDiagnosticFunctionManual.Text = config.textBoxDiagnosticFunctionManual_;

                BrushConverter bc = new BrushConverter();

                colorDefaultReadCell = (SolidColorBrush)bc.ConvertFromString(config.colorDefaultReadCell_);
                colorDefaultWriteCell = (SolidColorBrush)bc.ConvertFromString(config.colorDefaultWriteCell_); 
                colorErrorCell = (SolidColorBrush)bc.ConvertFromString(config.colorErrorCell_);

                labelColorCellRead.Background = colorDefaultReadCell;
                labelColorCellWrote.Background = colorDefaultWriteCell;
                labelColorCellError.Background = colorErrorCell;

                if(!(config.TextBoxPollingInterval_ is null))
                {
                    TextBoxPollingInterval.Text = config.TextBoxPollingInterval_;
                }

                if(config.CheckBoxSendValuesOnEditCoillsTable_ != null)
                {
                    CheckBoxSendValuesOnEditCoillsTable.IsChecked  = config.CheckBoxSendValuesOnEditCoillsTable_;
                }

                if(config.CheckBoxSendValuesOnEditHoldingTable_ != null)
                {
                    CheckBoxSendValuesOnEditHoldingTable.IsChecked = config.CheckBoxSendValuesOnEditHoldingTable_;
                }


                if (!statoConsole)
                {
                    var handle = GetConsoleWindow();
                    ShowWindow(handle, SW_HIDE);
                }

                // Scelgo quale richTextBox di status abilitare
                richTextBoxStatus.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
                richTextBoxStatus.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;

                Console.WriteLine("Caricata configurazione precedente\n");
            }
            catch
            {
                Console.WriteLine("Errore caricamento configurazione\n");
            }

            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/Template.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                TEMPLATE template = jss.Deserialize<TEMPLATE>(file_content);

                template_coilsOffset = 0;
                template_inputsOffset = 0;
                template_inputRegistersOffset = 0;
                template_HoldingOffset = 0;

                list_template_coilsTable = new ModBus_Item[template.dataGridViewCoils.Count()];
                list_template_inputsTable = new ModBus_Item[template.dataGridViewInput.Count()];
                list_template_inputRegistersTable = new ModBus_Item[template.dataGridViewInputRegister.Count()];
                list_template_holdingRegistersTable = new ModBus_Item[template.dataGridViewHolding.Count()];

                // Coils
                if (template.comboBoxCoilsOffset_ == "HEX")
                {
                    template_coilsOffset = int.Parse(template.textBoxCoilsOffset_, System.Globalization.NumberStyles.HexNumber);
                }
                else
                {
                    template_coilsOffset = int.Parse(template.textBoxCoilsOffset_);
                }

                // Inputs
                if (template.comboBoxInputOffset_ == "HEX")
                {
                    template_inputsOffset = int.Parse(template.textBoxInputOffset_, System.Globalization.NumberStyles.HexNumber);
                }
                else
                {
                    template_inputsOffset = int.Parse(template.textBoxInputOffset_);
                }

                // Input registers
                if (template.comboBoxInputRegOffset_ == "HEX")
                {
                    template_inputRegistersOffset = int.Parse(template.textBoxInputRegOffset_, System.Globalization.NumberStyles.HexNumber);
                }
                else
                {
                    template_inputRegistersOffset = int.Parse(template.textBoxInputRegOffset_);
                }

                // Holding registers
                if (template.comboBoxHoldingOffset_ == "HEX")
                {
                    template_HoldingOffset = int.Parse(template.textBoxHoldingOffset_, System.Globalization.NumberStyles.HexNumber);
                }
                else
                {
                    template_HoldingOffset = int.Parse(template.textBoxHoldingOffset_);
                }

                // Tabella coils
                for(int i = 0; i < template.dataGridViewCoils.Count(); i++) 
                {
                    if(template.comboBoxCoilsRegistri_ == "HEX")
                    {
                        template.dataGridViewCoils[i].Register = int.Parse(template.dataGridViewCoils[i].Register, System.Globalization.NumberStyles.HexNumber).ToString();
                    }

                    list_template_coilsTable[i] = template.dataGridViewCoils[i];
                }

                // Tabella inputs
                for (int i = 0; i < template.dataGridViewInput.Count(); i++)
                {
                    if (template.comboBoxCoilsRegistri_ == "HEX")
                    {
                        template.dataGridViewInput[i].Register = int.Parse(template.dataGridViewInput[i].Register, System.Globalization.NumberStyles.HexNumber).ToString();
                    }

                    list_template_inputsTable[i] = template.dataGridViewInput[i];
                }

                // Tabella input registers
                for (int i = 0; i < template.dataGridViewInputRegister.Count(); i++)
                {
                    if (template.comboBoxCoilsRegistri_ == "HEX")
                    {
                        template.dataGridViewInputRegister[i].Register = int.Parse(template.dataGridViewInputRegister[i].Register, System.Globalization.NumberStyles.HexNumber).ToString();
                    }

                    list_template_inputRegistersTable[i] = template.dataGridViewInputRegister[i];
                }

                // Tabella holdings
                for (int i = 0; i < template.dataGridViewHolding.Count(); i++)
                {
                    if (template.comboBoxCoilsRegistri_ == "HEX")
                    {
                        template.dataGridViewHolding[i].Register = int.Parse(template.dataGridViewHolding[i].Register, System.Globalization.NumberStyles.HexNumber).ToString();
                    }

                    list_template_holdingRegistersTable[i] = template.dataGridViewHolding[i];
                }
            }
            catch (Exception err)
            {
            Console.WriteLine("Errore caricamento configurazione\n");
            Console.WriteLine(err);
            }

        }

        private void salvaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione(true);
        }

        private void genera_tabelle_registri()
        {
            string content = "0";

            buttonReadCoils01.Dispatcher.Invoke((Action)delegate
            {
                // Disattivazione pulsanti fino al termine della generaizone delle tabelle
                buttonReadCoils01.IsEnabled = false;
                buttonReadCoilsRange.IsEnabled = false;
                buttonWriteCoils05.IsEnabled = false;
                buttonWriteCoils15.IsEnabled = false;
                buttonGoToCoilAddress.IsEnabled = false;

                buttonReadInput02.IsEnabled = false;
                buttonReadInputRange.IsEnabled = false;
                buttonGoToInputAddress.IsEnabled = false;

                buttonReadInputRegister04.IsEnabled = false;
                buttonReadInputRegisterRange.IsEnabled = false;
                buttonGoToInputRegisterAddress.IsEnabled = false;

                buttonReadHolding03.IsEnabled = false;
                buttonReadHoldingRange.IsEnabled = false;
                buttonWriteHolding06.IsEnabled = false;
                buttonWriteHolding16.IsEnabled = false;
                buttonGoToHoldingAddress.IsEnabled = false;

                Console.WriteLine("Generazione tabelle registri");

                list_coilsTable.Clear();
                list_inputsTable.Clear();
                list_inputRegistersTable.Clear();
                list_holdingRegistersTable.Clear();
            });

            ModBus_Item row = new ModBus_Item();

            checkBoxCreateTableAtBoot.Dispatcher.Invoke((Action)delegate {
            if ((bool)checkBoxCreateTableAtBoot.IsChecked)
            {
                MessageBox.Show("Generazione tabelle registri in corso", "Info");
                int stop = 10000;

                if (!(bool)checkBoxFollowModbusProtocol.IsChecked)
                    stop = 65536;

                for (int i = 1; i < stop; i++)
                {
                    // Tabella coils
                    row = new ModBus_Item();
                    row.Register = i.ToString();
                    row.Value = content;
                    row.ValueBin = i.ToString("X");

                    list_coilsTable.Add(row);


                    // Tabella input
                    row = new ModBus_Item();

                    if ((bool)checkBoxUseOffsetInTables.IsChecked)
                    {
                        row.Register = (i + 10000).ToString();
                        row.ValueBin = (i + 10000).ToString("X");
                    }
                    else
                    {
                        row.Register = i.ToString();
                        row.ValueBin = (i).ToString("X");
                    }

                    row.Value = content;
                    list_inputsTable.Add(row);

                    // Tabella input register
                    row = new ModBus_Item();

                    if ((bool)checkBoxUseOffsetInTables.IsChecked)
                    {
                        row.Register = (i + 30000).ToString();
                        row.ValueBin = (i + 30000).ToString("X");
                    }
                    else
                    {
                        row.Register = i.ToString();
                        row.ValueBin = (i).ToString("X");
                    }

                    row.Value = content;
                    list_inputRegistersTable.Add(row);

                    // Tabella holding register
                    row = new ModBus_Item();

                    if ((bool)checkBoxUseOffsetInTables.IsChecked)
                    {
                        row.Register = (i + 40000).ToString();
                        row.ValueBin = (i + 40000).ToString("X");
                    }
                    else
                    {
                        row.Register = i.ToString();
                        row.ValueBin = (i).ToString("X");
                    }

                    row.Value = content;
                    list_holdingRegistersTable.Add(row);

                    // Tabella summary   (eliminata)
                    //row = new DataGridViewRow();
                    //row.CreateCells(dataGridViewSummary);
                    //row.Cells[0].Value = i;
                    //row.Cells[1].Value = content;
                    //row.Cells[2].Value = i + 10000;
                    //row.Cells[3].Value = content;
                    //row.Cells[4].Value = i + 30000;
                    //row.Cells[5].Value = content;
                    //row.Cells[6].Value = i + 40000;
                    //row.Cells[7].Value = content;
                    //dataGridViewSummary.Rows.Add(row);

                    // DEBUG
                    //Console.WriteLine(i + ToString());

                    if ((i + 1) % (stop / 20) == 0)
                    {
                        Console.WriteLine(((i + 1) / (stop / 100)).ToString() + "% Completato");
                    }
                }
                }
            });

            Console.WriteLine("Operazione terminata\n");

            buttonReadCoils01.Dispatcher.Invoke((Action)delegate
            {
                // Attivazione pulsanti al termine della generaizone delle tabelle
                buttonReadCoils01.IsEnabled = true;
                buttonReadCoilsRange.IsEnabled = true;
                buttonWriteCoils05.IsEnabled = true;
                buttonWriteCoils15.IsEnabled = true;
                buttonGoToCoilAddress.IsEnabled = true;

                buttonReadInput02.IsEnabled = true;
                buttonReadInputRange.IsEnabled = true;
                buttonGoToInputAddress.IsEnabled = true;

                buttonReadInputRegister04.IsEnabled = true;
                buttonReadInputRegisterRange.IsEnabled = true;
                buttonGoToInputRegisterAddress.IsEnabled = true;

                buttonReadHolding03.IsEnabled = true;
                buttonReadHoldingRange.IsEnabled = true;
                buttonWriteHolding06.IsEnabled = true;
                buttonWriteHolding16.IsEnabled = true;
                buttonGoToHoldingAddress.IsEnabled = true;


                // Dopo 2.5s secondi chiudo la console

                if ((bool)checkBoxCloseConsolAfterBoot.IsChecked)
                {
                    var handle = GetConsoleWindow();
                    ShowWindow(handle, SW_HIDE);
                }
            });
            Thread.Sleep(500);
            Thread.CurrentThread.Abort();
        }

        private void buttonTcpActive_Click(object sender, RoutedEventArgs e)
        {
            String ip_address = "";
            String port = "";

            ip_address = textBoxTcpClientIpAddress.Text;
            port = textBoxTcpClientPort.Text;


            try
            {
                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                client.Close();

                radioButtonModeSerial.IsEnabled = false;
                radioButtonModeTcp.IsEnabled = false;

                //richTextBoxAppend(richTextBoxStatus, "Connected to " + ip_address + ":" + port); ;

            }
            catch
            {
                Console.WriteLine("Impossibile stabilire una connessione con il server");
                richTextBoxAppend(richTextBoxStatus, "Failed to connect to " + ip_address + ":" + port); ;

                return;
            }



            if (pictureBoxTcp.Background == Brushes.LightGray)
            {
                // Tcp attivo
                pictureBoxTcp.Background = Brushes.Lime;
                pictureBoxRunningAs.Background = Brushes.Lime;

                ModBus = new ModBus_Chicco(serialPort, textBoxTcpClientIpAddress.Text, textBoxTcpClientPort.Text, "TCP", pictureBoxIsResponding, pictureBoxIsSending, richTextBoxOutgoingPackets, richTextBoxIncomingPackets);
                ModBus.open();

                richTextBoxAppend(richTextBoxStatus, "Connected to " + ip_address + ":" + port); ;

                buttonTcpActive.Content = "Disconnect";
                menuItemToolBit.IsEnabled = true;
                menuItemToolWord.IsEnabled = true;
                menuItemToolByte.IsEnabled = true;
                salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = false;
                caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = false;
                gestisciDatabaseToolStripMenuItem.IsEnabled = false;

                textBoxTcpClientIpAddress.IsEnabled = false;
                textBoxTcpClientPort.IsEnabled = false;
            }
            else
            {
                // Tcp disattivo
                pictureBoxTcp.Background = Brushes.LightGray;
                pictureBoxRunningAs.Background = Brushes.LightGray;

                richTextBoxAppend(richTextBoxStatus, "Disconnected");

                buttonTcpActive.Content = "Connect";
                menuItemToolBit.IsEnabled = false;
                menuItemToolWord.IsEnabled = false;
                menuItemToolByte.IsEnabled = false;
                salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = true;
                caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = true;
                gestisciDatabaseToolStripMenuItem.IsEnabled = true;

                radioButtonModeSerial.IsEnabled = true;
                radioButtonModeTcp.IsEnabled = true;

                textBoxTcpClientIpAddress.IsEnabled = true;
                textBoxTcpClientPort.IsEnabled = true;
            }
        }

        //----------------------------------------------------------------------------------
        //-------------------------------Comandi force coil---------------------------------
        //----------------------------------------------------------------------------------

        private void buttonReadCoils01_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset) + P.uint_parser(textBoxCoilsAddress01, comboBoxCoilsAddress01);

                if (uint.Parse(textBoxCoilNumber.Text) > 123)
                {
                    MessageBox.Show("Numero di registri troppo elevato", "Info");
                }
                else
                {
                    String[] response = ModBus.readCoilStatus_01(byte.Parse(textBoxModbusAddress.Text), address_start, uint.Parse(textBoxCoilNumber.Text));

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        // Uso le righe esistenti
                        updateRowTable(list_coilsTable, null, address_start, response, colorDefaultReadCell);
                    }
                    else
                    {
                        // Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_coilsTable, null, address_start, response, colorDefaultReadCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                        else
                            insertRowsTable(list_coilsTable, null, address_start, response, colorDefaultReadCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                    }

                    applyTemplateCoils();
                }
            }
            catch { }
        }

        private void buttonReadCoilsRange_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset) + P.uint_parser(textBoxCoilsRange_A, comboBoxCoilsRange_A);
                uint coil_len = P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset) + P.uint_parser(textBoxCoilsRange_B, comboBoxCoilsRange_B) - address_start + 1;

                uint repeatQuery = coil_len / 120;

                if (coil_len % 120 != 0)
                {
                    repeatQuery += 1;
                }

                String[] response = new string[coil_len];

                for (int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1))
                    {
                        Array.Copy(ModBus.readCoilStatus_01(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), coil_len % 120), 0, response, 120 * i, coil_len % 120);
                    }
                    else
                    {
                        Array.Copy(ModBus.readCoilStatus_01(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), 120), 0, response, 120 * i, 120);
                    }
                }

                if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                {
                    // Uso le righe esistenti
                    updateRowTable(list_coilsTable, null, address_start, response, colorDefaultReadCell);
                }
                else
                {
                    // Cancello la tabella e inserisco le nuove righe
                    if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                        insertRowsTable(list_coilsTable, null, address_start - P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset), response, colorDefaultReadCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                    else
                        insertRowsTable(list_coilsTable, null, address_start, response, colorDefaultReadCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                }
            }
            catch { }
        }

        private void buttonWriteCoils05_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset) + P.uint_parser(textBoxCoilsAddress05, comboBoxCoilsAddress05);

                if (ModBus.forceSingleCoil_05(byte.Parse(textBoxModbusAddress.Text), address_start, uint.Parse(textBoxCoilsValue05.Text)))
                {
                    String[] value = { textBoxCoilsValue05.Text };

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        // Uso le righe esistenti
                        updateRowTable(list_coilsTable, null, address_start, value, colorDefaultWriteCell);
                    }
                    else
                    {
                        // Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_coilsTable, null, address_start - P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset), value, colorDefaultWriteCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                        else
                            insertRowsTable(list_coilsTable, null, address_start, value, colorDefaultWriteCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                    }
                }
                else
                {
                    MessageBox.Show("Errore nel settaggio del registro", "Alert");

                    list_coilsTable[(int)(address_start)].Color = Brushes.Red.ToString();
                }
            }
            catch { }
        }

        private void buttonWriteCoils15_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxCoilsAddress15_A, comboBoxCoilsAddress15_A);

                if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                uint address_stop = P.uint_parser(textBoxCoilsOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxCoilsAddress15_B, comboBoxCoilsAddress15_B);

                if (address_stop > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_stop = address_stop - 40001;
                }

                bool[] buffer = new bool[address_stop - address_start + 1];

                for (int i = 0; i < (address_stop - address_start + 1); i++)
                {
                    buffer[i] = uint.Parse(textBoxCoilsValue15.Text) > 0;
                }

                if (ModBus.forceMultipleCoils_15(byte.Parse(textBoxModbusAddress.Text), address_start, buffer))
                {
                    String[] value = new String[address_stop - address_start + 1];

                    for (int i = 0; i < (address_stop - address_start + 1); i++)
                    {
                        value[i] = uint.Parse(textBoxCoilsValue15.Text) > 0 ? "1" : "0";
                    }

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        // Uso le righe esistenti
                        updateRowTable(list_coilsTable, null, address_start, value, colorDefaultWriteCell);
                    }
                    else
                    {
                        // Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_coilsTable, null, address_start - P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset), value, colorDefaultWriteCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], null);
                        else
                            insertRowsTable(list_coilsTable, null, address_start, value, colorDefaultWriteCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], null);
                    }
                }
                else
                {
                    MessageBox.Show("Error setting coils values", "Alert");

                    list_coilsTable[(int)(address_start)].Color = Brushes.Red.ToString();
                }
            }
            catch { }
        }

        private void buttonGoToCoilAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToCoilAddress.Text);

            //dataGridViewCoils.Scroo = dataGridViewCoils.Rows[index].Cells[0];
        }

        //----------------------------------------------------------------------------------
        //--------------------------------Comandi read input--------------------------------
        //----------------------------------------------------------------------------------


        private void buttonReadInput02_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxInputOffset, comboBoxInputOffset) + P.uint_parser(textBoxInputAddress02, comboBoxInputAddress02);

                if (uint.Parse(textBoxInputNumber.Text) > 123)
                {
                    MessageBox.Show("Numero di registri troppo elevato", "Info");
                }
                else
                {
                    String[] response = ModBus.readInputStatus_02(byte.Parse(textBoxModbusAddress.Text), address_start, uint.Parse(textBoxInputNumber.Text));

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        // Uso le righe esistenti
                        updateRowTable(list_inputsTable, null, address_start, response, colorDefaultReadCell);
                    }
                    else
                    {
                        // Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_inputsTable, null, address_start - P.uint_parser(textBoxInputOffset, comboBoxInputOffset), response, colorDefaultReadCell, comboBoxInputRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                        else
                            insertRowsTable(list_inputsTable, null, address_start, response, colorDefaultReadCell, comboBoxInputRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                    }

                    applyTemplateInputs();
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }
        }

        private void buttonReadInputRange_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxInputOffset, comboBoxInputOffset) + P.uint_parser(textBoxInputRange_A, comboBoxInputRange_A);
                uint input_len = P.uint_parser(textBoxInputOffset, comboBoxInputOffset) + P.uint_parser(textBoxInputRange_B, comboBoxInputRange_B) - address_start + 1;

                // Anche se le coils le leggo a bit e non byte, tengo 120 coils per lettura per ora, prox release si può incrementare
                uint repeatQuery = input_len / 120;

                if (input_len % 120 != 0)
                {
                    repeatQuery += 1;
                }

                String[] response = new string[input_len];

                for (int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1))
                    {
                        Array.Copy(ModBus.readInputStatus_02(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), input_len % 120), 0, response, 120 * i, input_len % 120);
                    }
                    else
                    {
                        Array.Copy(ModBus.readInputStatus_02(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), 120), 0, response, 120 * i, 120);
                    }
                }

                if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                {
                    // Uso le righe esistenti
                    updateRowTable(list_inputsTable, null, address_start, response, colorDefaultReadCell);
                }
                else
                {
                    // Cancello la tabella e inserisco le nuove righe
                    if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                        insertRowsTable(list_inputsTable, null, address_start - P.uint_parser(textBoxInputOffset, comboBoxInputOffset), response, colorDefaultReadCell, comboBoxInputRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                    else
                        insertRowsTable(list_inputsTable, null, address_start, response, colorDefaultReadCell, comboBoxInputRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                }

                applyTemplateInputs();
            }
            catch { }
        }

        private void buttonGoToInputAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToInputAddress.Text);

            if (index > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)
                index = index - 10001;

            //dataGridViewInput.FirstDisplayedCell = dataGridViewInput.Rows[index].Cells[0];
        }

        // ----------------------------------------------------------------------------------
        // ---------------------------Comandi input register---------------------------------
        // ----------------------------------------------------------------------------------

        // Read input register FC04
        private void buttonReadInputRegister04_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                uint address_start = P.uint_parser(textBoxInputRegOffset, comboBoxInputRegOffset) + P.uint_parser(textBoxInputRegisterAddress04, comboBoxInputRegisterAddress04);

                if (uint.Parse(textBoxInputRegisterNumber.Text) > 123)
                {
                    MessageBox.Show("Numero di registri troppo elevato", "Info");
                }
                else
                {
                    if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                    {
                        address_start = address_start - 30001;
                    }

                    String[] response = ModBus.readInputRegister_04(byte.Parse(textBoxModbusAddress.Text), address_start, uint.Parse(textBoxInputRegisterNumber.Text));

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        //Uso le righe esistenti
                        updateRowTable(list_inputRegistersTable, null, address_start, response, colorDefaultReadCell);
                    }
                    else
                    {
                        //Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_inputRegistersTable, null, address_start - P.uint_parser(textBoxInputRegOffset, comboBoxInputRegOffset), response, colorDefaultReadCell, comboBoxInputRegRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxInputRegValori.SelectedValue.ToString().Split(' ')[1]);
                        else
                            insertRowsTable(list_inputRegistersTable, null, address_start, response, colorDefaultReadCell, comboBoxInputRegRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxInputRegValori.SelectedValue.ToString().Split(' ')[1]);
                    }

                    applyTemplateInputRegister();
                }
            }
            catch { }
        }

        // Read input register range
        private void buttonReadInputRegisterRange_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxInputRegOffset, comboBoxInputRegOffset) + P.uint_parser(textBoxInputRegisterRange_A, comboBoxInputRegisterRange_A);
                uint register_len = P.uint_parser(textBoxInputRegOffset, comboBoxInputRegOffset) + P.uint_parser(textBoxInputRegisterRange_B, comboBoxInputRegisterRange_B) - address_start + 1;

                uint repeatQuery = register_len / 120;

                if (register_len % 120 != 0)
                {
                    repeatQuery += 1;
                }

                String[] response = new string[register_len];

                for (int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1))
                    {
                        Array.Copy(ModBus.readInputRegister_04(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), register_len % 120), 0, response, 120 * i, register_len % 120);
                    }
                    else
                    {
                        Array.Copy(ModBus.readInputRegister_04(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), 120), 0, response, 120 * i, 120);
                    }
                }

                if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 30001;
                }

                if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                {
                    //Uso le righe esistenti
                    updateRowTable(list_inputRegistersTable, null, address_start, response, colorDefaultReadCell);
                }
                else
                {
                    //Cancello la tabella e inserisco le nuove righe
                    if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                        insertRowsTable(list_inputRegistersTable, null, address_start - P.uint_parser(textBoxInputRegOffset, comboBoxInputRegOffset), response, colorDefaultReadCell, comboBoxInputRegRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxInputRegValori.SelectedValue.ToString().Split(' ')[1]);
                    else
                        insertRowsTable(list_inputRegistersTable, null, address_start, response, colorDefaultReadCell, comboBoxInputRegRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxInputRegValori.SelectedValue.ToString().Split(' ')[1]);
                }

                applyTemplateInputRegister();
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }
        }

        //Go to input register
        private void buttonGoToInputRegisterAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToInputRegisterAddress.Text);

            if (index > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)
                index = index - 30001;

            //dataGridViewInputRegister.FirstDisplayedCell = dataGridViewInputRegister.Rows[index].Cells[0];
        }

        // ----------------------------------------------------------------------------------
        // ---------------------------Comandi holding register-------------------------------
        // ----------------------------------------------------------------------------------

        // Read holding register
        private void buttonReadHolding03_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingAddress03, comboBoxHoldingAddress03);

                if (uint.Parse(textBoxHoldingRegisterNumber.Text) > 123)
                {
                    MessageBox.Show("Numero di registri troppo elevato", "Info");
                }
                else
                {
                    if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    // Se indirizzo espresso in 30001+ imposto offset a 0
                    {
                        address_start = address_start - 40001;
                    }

                    string[] response = ModBus.readHoldingRegister_03(byte.Parse(textBoxModbusAddress.Text), address_start, uint.Parse(textBoxHoldingRegisterNumber.Text));

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        // Uso le righe esistenti
                        updateRowTable(list_holdingRegistersTable, null, address_start, response, colorDefaultReadCell);
                    }
                    else
                    {
                        // Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_holdingRegistersTable, null, address_start - P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset), response, colorDefaultReadCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                        else
                            insertRowsTable(list_holdingRegistersTable, null, address_start, response, colorDefaultReadCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                    }

                    // Applico le note ai registri
                    applyTemplateHoldingRegister();
                }
            }
            catch { }
        }

        public void applyTemplateCoils()
        {
            // Carico le etichette dal template per la tabella corrente
            System.Globalization.NumberStyles registerFormat = comboBoxCoilsRegistri.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;
            System.Globalization.NumberStyles offsetFormat = comboBoxCoilsOffset.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;

            int offsetValue = int.Parse(textBoxCoilsOffset.Text, offsetFormat); // Offset utente input

            // Passo fuori ogni riga della tabella dei registri
            for (int a = 0; a < list_coilsTable.Count(); a++)
            {
                // Cerco una corrispondenza nel file template
                for (int i = 0; i < list_template_coilsTable.Count(); i++)
                {
                    // Se trovo una corrispondenza esco dal for (list_template_inputsTable.Register,template_inputsOffset,offsetValue sono già in DEC, list_inputsTable[a].Register dipende DEC o HEX)
                    if ((int.Parse(list_template_coilsTable[i].Register) + template_coilsOffset) == (int.Parse(list_coilsTable[a].Register, registerFormat) + offsetValue))
                    {
                        list_coilsTable[a].Notes = list_template_coilsTable[i].Notes;
                        break;
                    }
                }
            }
        }

        public void applyTemplateInputs()
        {
            // Carico le etichette dal template per la tabella corrente
            System.Globalization.NumberStyles registerFormat = comboBoxInputRegistri.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;
            System.Globalization.NumberStyles offsetFormat = comboBoxInputOffset.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;

            int offsetValue = int.Parse(textBoxInputOffset.Text, offsetFormat); // Offset utente input

            // Passo fuori ogni riga della tabella dei registri
            for (int a = 0; a < list_inputsTable.Count(); a++)
            {
                // Cerco una corrispondenza nel file template
                for (int i = 0; i < list_template_inputsTable.Count(); i++)
                {
                    // Se trovo una corrispondenza esco dal for (list_template_inputsTable.Register,template_inputsOffset,offsetValue sono già in DEC, list_inputsTable[a].Register dipende DEC o HEX)
                    if ((int.Parse(list_template_inputsTable[i].Register) + template_inputsOffset) == (int.Parse(list_inputsTable[a].Register, registerFormat) + offsetValue))
                    {
                        list_inputsTable[a].Notes = list_template_inputsTable[i].Notes;
                        break;
                    }
                }
            }
        }

        public void applyTemplateInputRegister()
        {
            // Carico le etichette dal template per la tabella corrente
            System.Globalization.NumberStyles registerFormat = comboBoxInputRegRegistri.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;
            System.Globalization.NumberStyles offsetFormat = comboBoxInputRegOffset.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;
            
            int offsetValue = int.Parse(textBoxInputRegOffset.Text, offsetFormat);

            // Passo fuori ogni riga della tabella dei registri
            for (int a = 0; a < list_inputRegistersTable.Count(); a++)
            {
                // Cerco una corrispondenza nel file template
                for (int i = 0; i < list_template_inputRegistersTable.Count(); i++)
                {
                    // Se trovo una corrispondenza esco dal for (list_template_inputRegistersTable.Register,template_inputRegistersOffset,offsetValue sono già in DEC, list_inputRegistersTable[a].Register dipende DEC o HEX)
                    if ((int.Parse(list_template_inputRegistersTable[i].Register) + template_inputRegistersOffset) == (int.Parse(list_inputRegistersTable[a].Register, registerFormat) + offsetValue))
                    {
                        list_inputRegistersTable[a].Notes = list_template_inputRegistersTable[i].Notes;
                        break;
                    }
                }
            }
        }

        public void applyTemplateHoldingRegister()
        {
            // Carico le etichette dal template per la tabella corrente
            System.Globalization.NumberStyles registerFormat = comboBoxHoldingRegistri.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;
            System.Globalization.NumberStyles offsetFormat = comboBoxHoldingOffset.SelectedItem.ToString().Split(' ')[1] == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;
            
            int offsetValue = int.Parse(textBoxHoldingOffset.Text, offsetFormat);

            // Passo fuori ogni riga della tabella dei registri
            for (int a = 0; a < list_holdingRegistersTable.Count(); a++)
            {
                // Cerco una corrispondenza nel file template
                for (int i = 0; i < list_template_holdingRegistersTable.Count(); i++)
                {
                    // Se trovo una corrispondenza esco dal for (list_template_inputRegistersTable.Register,template_inputRegistersOffset,offsetValue sono già in DEC, list_inputRegistersTable[a].Register dipende DEC o HEX)
                    if ((int.Parse(list_template_holdingRegistersTable[i].Register) + template_HoldingOffset) == (int.Parse(list_holdingRegistersTable[a].Register, registerFormat) + offsetValue))
                    {
                        list_holdingRegistersTable[a].Notes = list_template_holdingRegistersTable[i].Notes;
                        break;
                    }
                }
            }
        }


        // Preset single register
        private void buttonWriteHolding06_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingAddress06, comboBoxHoldingAddress06);

                if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                if (ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress.Text), address_start, P.uint_parser(textBoxHoldingValue06, comboBoxHoldingValue06)))
                {
                    String[] value = { P.uint_parser(textBoxHoldingValue06, comboBoxHoldingValue06).ToString() };
                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        //Uso le righe esistenti
                        updateRowTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell);
                    }
                    else
                    {
                        //Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_holdingRegistersTable, null, address_start - P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset), value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                        else
                            insertRowsTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                    }
                }
                else
                {
                    MessageBox.Show("Errore nel settaggio del registro", "Alert");

                    list_holdingRegistersTable[(int)(address_start)].Color = Brushes.Red.ToString();
                }
            }
            catch { }

        }

        //Preset multiple register
        private void buttonWriteHolding16_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingAddress16_A, comboBoxHoldingAddress16_A);

                if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                uint address_stop = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingAddress16_B, comboBoxHoldingAddress16_B);

                if (address_stop > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_stop = address_stop - 40001;
                }

                uint[] buffer = new uint[address_stop - address_start + 1];

                for(int i = 0; i < (address_stop - address_start + 1); i++)
                {
                    buffer[i] = P.uint_parser(textBoxHoldingValue16, comboBoxHoldingValue16);
                }

                if (ModBus.presetMultipleRegisters_16(byte.Parse(textBoxModbusAddress.Text), address_start, buffer))
                {
                    String[] value = new String[address_stop - address_start + 1];

                    for (int i = 0; i < (address_stop - address_start + 1); i++)
                    {
                        value[i] = P.uint_parser(textBoxHoldingValue16, comboBoxHoldingValue16).ToString();
                    }

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        //Uso le righe esistenti
                        updateRowTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell);
                    }
                    else
                    {
                        //Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_holdingRegistersTable, null, address_start - P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset), value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                        else
                            insertRowsTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                    }
                }
                else
                {
                    MessageBox.Show("Errore nel settaggio del registro", "Alert");

                    list_holdingRegistersTable[(int)(address_start)].Color = Brushes.Red.ToString();
                }
            }
            catch { }
        }

        //Read holding register range
        private void buttonReadHoldingRange_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingRange_A, comboBoxHoldingRange_A);
                uint register_len = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingRange_B, comboBoxHoldingRange_B) - address_start + 1;

                uint repeatQuery = register_len / 120;

                if(register_len % 120 != 0)
                {
                    repeatQuery += 1;
                }

                String[] response = new string[register_len];

                for(int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1))
                    {
                        Array.Copy(ModBus.readHoldingRegister_03(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), register_len % 120), 0, response, 120 * i, register_len % 120);
                    }
                    else
                    {
                        Array.Copy(ModBus.readHoldingRegister_03(byte.Parse(textBoxModbusAddress.Text), address_start + (uint)(120 * i), 120), 0, response, 120 * i, 120);
                    }
                }

                if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                {
                    //Uso le righe esistenti
                    updateRowTable(list_holdingRegistersTable, null, address_start, response, colorDefaultReadCell);
                }
                else
                {
                    //Cancello la tabella e inserisco le nuove righe
                    if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                        insertRowsTable(list_holdingRegistersTable, null, address_start - P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset), response, colorDefaultReadCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                    else
                        insertRowsTable(list_holdingRegistersTable, null, address_start, response, colorDefaultReadCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                }

                applyTemplateHoldingRegister();
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }
        }

        //Go to holding register
        private void buttonGoToHoldingAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToHoldingAddress.Text);

            if (index > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)
                index = index - 40001;

            //dataGridViewHolding.FirstDisplayedCell = dataGridViewHolding.Rows[index].Cells[0];
        }


        //Altri pulsanti nella grafica
        private void guidaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\\Manuali\\Guida_ModBus_Client.pdf");
            }
            catch
            {
                MessageBox.Show("Ancora da scrivere :-)", "Hey");
            }
        }

        private void infoToolStripMenuItem1_Click(object sender, RoutedEventArgs e)
        {
            Info info = new Info(title, version);

            info.Show();
        }

        private void buttonClearSent_Click(object sender, RoutedEventArgs e)
        {
            richTextBoxOutgoingPackets.Document.Blocks.Clear();
            richTextBoxOutgoingPackets.AppendText("\n");
        }

        private void buttonClearReceived_Click(object sender, RoutedEventArgs e)
        {
            richTextBoxIncomingPackets.Document.Blocks.Clear();
            richTextBoxIncomingPackets.AppendText("\n");
        }

        private void richTextBoxAppend(RichTextBox richTextBox, String append)
        {
            richTextBox.AppendText(DateTime.Now.ToString("hh:MM:ss") + " " + append + "\n");

        }

        private void buttonClearSerialStatus_Click(object sender, RoutedEventArgs e)
        {
            richTextBoxStatus.Document.Blocks.Clear();
            richTextBoxStatus.AppendText("\n");
        }

        private void esciToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione(false);
            this.Close();
        }

        private void impostazioniToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            tabControlMain.SelectedIndex = 7;
        }

        //Inserisce le righe nella tabella
        public void insertRowsTable(ObservableCollection<ModBus_Item> tab_1, ObservableCollection<ModBus_Item> tab_2, uint address_start, String[] response, SolidColorBrush cellBackGround, String formatRegister, String formatVal)
        {
            tab_1.Clear();

            if (response != null)
            {
                for (int i = 0; i < response.Length; i++)
                {
                    //Tabella 1
                    ModBus_Item row = new ModBus_Item();

                    // Cella con numero del registro
                    if (formatRegister == "DEC" || formatRegister == null)
                    {
                        row.Register = (address_start + i).ToString();
                    }
                    else
                    {
                        row.Register = "0x" + (address_start + i).ToString("X").PadLeft(4, '0');
                    }

                    // Cella con valore letto
                    if (formatVal == "DEC" || formatVal == null)
                    {
                        row.Value = (response[i]).ToString();
                    }
                    else
                    {
                        row.Value = "0x" + (int.Parse(response[i])).ToString("X").PadLeft(4, '0');
                    }

                    //Colorazione celle
                    if ((bool)checkBoxCellColorMode.IsChecked)
                    {
                        if (int.Parse(response[i]) > 0)
                        {
                            row.Color = cellBackGround.ToString();
                        }
                    }
                    else
                    {
                        if (i % 2 == 0)
                        {
                            row.Color = cellBackGround.ToString();
                        }
                    }

                    // Il valore in binario lo metto sempre tanto poi nelle coils ed inputs è nascosto
                    row.ValueBin = Convert.ToString(UInt16.Parse(response[i]) >> 8, 2).PadLeft(8, '0') + " " + Convert.ToString((UInt16)(UInt16.Parse(response[i]) << 8) >> 8, 2).PadLeft(8, '0');

                    /*
                    //Quarta colonna tabelle (se > 3 la seconda colonna contiene il valore letto in hex
                    //altrimenti se == 3 il valore dell'address dell'indirizzo in hex)
                    if (row.Cells.Count > 3)
                    {
                        //row.Cells[2].Value = (int.Parse(response[i])).ToString("B");
                        row.Cells[2].Value = Convert.ToString(int.Parse(response[i])>>8,2).PadLeft(8,'0') + " " + Convert.ToString(byte.Parse(response[i]),2).PadLeft(8, '0');
                        if (i % 2 == 0)
                            row.Cells[2].Style.Background = cellBackGround;
                        //Quarta colonna
                        row.Cells[3].Value = (int.Parse(response[i])).ToString("X");
                        if (i % 2 == 0)
                            row.Cells[3].Style.Background = cellBackGround;

                        //Quinta colonna
                        row.Cells[4].Value = (address_start + i).ToString("X");
                    }
                    else if (row.Cells.Count == 3)
                    {
                        //Quarta colonna (Valore in hex)
                        row.Cells[2].Value = (address_start + i).ToString("X");
                        if (i % 2 == 0)
                            row.Cells[2].Style.Background = cellBackGround;
                    }

                    //Quinta colonna (Numero registro in hex)
                    /*if (row.Cells.Count > 4)
                    {
                        row.Cells[4].Value = (address_start + i).ToString("X");
                        //row.Cells[4].Style.Background = cellBackGround;
                    }*/
                    tab_1.Add(row);
                }

                if (tab_2 != null)
                {
                    for (int i = 0; i < response.Length; i++)
                    {
                        tab_2.Clear();

                        //Tabella 1
                        ModBus_Item row = new ModBus_Item();

                        row.Register = (address_start + i).ToString();
                        row.Value = response[i];
                        row.Color = cellBackGround.ToString();
                        row.ValueBin = (int.Parse(response[i])).ToString("X");

                        //row.Cells[3].Value = (address_start + i).ToString("X");
                        //row.Cells[3].Style.Background = cellBackGround;

                        tab_2.Add(row);
                    }
                }
            }
        }

        // Aggiorna riga nella tabella che esiste gia'
        public void updateRowTable(ObservableCollection<ModBus_Item> tab_1, ObservableCollection<ModBus_Item> tab_2, uint address_start, String[] response, SolidColorBrush cellBackGround)
        {
            if (response != null)
            {
                for (int i = 0; i < response.Length; i++)
                {
                    tab_1[(int)(address_start) + i].Value = response[i];
                    tab_1[(int)(address_start) + i].Color = cellBackGround.ToString();

                    tab_1[(int)(address_start) + i].ValueBin = (int.Parse(response[i])).ToString("X");

                    if (tab_2 != null)
                    {
                        tab_2[(int)(address_start) + i].Value = response[i];
                        tab_2[(int)(address_start) + i].Color = cellBackGround.ToString();

                        tab_2[(int)(address_start) + i].ValueBin = (int.Parse(response[i])).ToString("X");
                    }
                }
            }
        }



        private void buttonSendDiagnosticQuery_Click(object sender, RoutedEventArgs e)
        {
            byte[] diagnostic_codes = { 0x00, 0x01, 0x02, 0x03, 0x04, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F, 0x10, 0x11, 0x12, 0x14 };

            //DEBUG
            //MessageBox.Show(comboBoxDiagnosticFunction.SelectedIndex.ToString());
            //SelectedIndex considera -1 la cella vuota e 0 il primo elemento del menu a tendina

            if (comboBoxDiagnosticFunction.SelectedIndex >= 0)
            {
                try
                {
                    textBoxDiagnosticResponse.Text = ModBus.diagnostics_08(byte.Parse(textBoxModbusAddress.Text), diagnostic_codes[comboBoxDiagnosticFunction.SelectedIndex], UInt16.Parse(textBoxDiagnosticData.Text));
                }
                catch
                {
                    textBoxDiagnosticResponse.Text = "Error executing command";
                }
            }
            else
            {
                MessageBox.Show("Valore scelto non valido", "Alert");
            }
        }

        private void buttonColorCellRead_Click(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                colorDefaultReadCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellRead.Background = colorDefaultReadCell;
            }
        }

        private void buttonCellWrote_Click(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                colorDefaultWriteCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellWrote.Background = colorDefaultWriteCell;
            }
        }

        private void buttonColorCellError_Click(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                colorErrorCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellError.Background = colorErrorCell;
            }
        }

        // Salvataggio pacchetti inviati
        private void buttonExportSentPackets_Click(object sender, RoutedEventArgs e)
        {
            saveFileDialogBox = new SaveFileDialog();

            saveFileDialogBox.DefaultExt = ".txt";
            saveFileDialogBox.AddExtension = false;
            saveFileDialogBox.FileName = "Pacchetti_inviati_" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "." + DateTime.Now.Month.ToString().PadLeft(2, '0') + "." + DateTime.Now.Year.ToString().PadLeft(2, '0');
            saveFileDialogBox.Filter = "Text|*.txt|Log|*.log";
            saveFileDialogBox.Title = "Salva Log pacchetti inviati";
            saveFileDialogBox.ShowDialog();

            if (saveFileDialogBox.FileName != "")
            {
                StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                TextRange textRange = new TextRange(
                    // TextPointer to the start of content in the RichTextBox.
                    richTextBoxIncomingPackets.Document.ContentStart,
                    // TextPointer to the end of content in the RichTextBox.
                    richTextBoxIncomingPackets.Document.ContentEnd
                );


                writer.Write(textRange.Text);
                writer.Dispose();
                writer.Close();
            }
        }

        // Salvataggio pacchetti ricevuti
        private void buttonExportReceivedPackets_Click(object sender, RoutedEventArgs e)
        {
            saveFileDialogBox = new SaveFileDialog();

            saveFileDialogBox.DefaultExt = ".txt";
            saveFileDialogBox.AddExtension = false;
            saveFileDialogBox.FileName = "Pacchetti_ricevuti_" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "." + DateTime.Now.Month.ToString().PadLeft(2, '0') + "." + DateTime.Now.Year.ToString().PadLeft(2, '0');
            saveFileDialogBox.Filter = "Text|*.txt|Log|*.log";
            saveFileDialogBox.Title = "Salva Log pacchetti inviati";
            saveFileDialogBox.ShowDialog();

            if (saveFileDialogBox.FileName != "")
            {
                StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                TextRange textRange = new TextRange(
                    // TextPointer to the start of content in the RichTextBox.
                    richTextBoxIncomingPackets.Document.ContentStart,
                    // TextPointer to the end of content in the RichTextBox.
                    richTextBoxIncomingPackets.Document.ContentEnd
                );


                writer.Write(textRange.Text);
                writer.Dispose();
                writer.Close();
            }
        }

        private void buttonWriteHolding06_b_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(textBoxHoldingAddress06_b, comboBoxHoldingAddress06_b);

                if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                if (ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress.Text), address_start, P.uint_parser(textBoxHoldingValue06_b, comboBoxHoldingValue06_b)))
                {
                    String[] value = { P.uint_parser(textBoxHoldingValue06_b, comboBoxHoldingValue06_b).ToString() };

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        //Uso le righe esistenti
                        updateRowTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell);
                    }
                    else
                    {
                        //Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_holdingRegistersTable, null, address_start - P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset), value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                        else
                            insertRowsTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                    }
                }
                else
                {
                    MessageBox.Show("Errore nel settaggio del registro", "Alert");

                    list_holdingRegistersTable[(int)(address_start)].Color = Brushes.Red.ToString();
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }
        }

        private void buttonWriteCoils05_B_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset) + P.uint_parser(textBoxCoilsAddress05_b, comboBoxCoilsAddress05_b);

                if (ModBus.forceSingleCoil_05(byte.Parse(textBoxModbusAddress.Text), address_start, uint.Parse(textBoxCoilsValue05_b.Text)))
                {
                    String[] value = { textBoxCoilsValue05_b.Text };

                    if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                    {
                        //Uso le righe esistenti
                        updateRowTable(list_coilsTable, null, address_start, value, colorDefaultWriteCell);
                    }
                    else
                    {
                        //Cancello la tabella e inserisco le nuove righe
                        if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                            insertRowsTable(list_coilsTable, null, address_start - P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset), value, colorDefaultWriteCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                        else
                            insertRowsTable(list_coilsTable, null, address_start, value, colorDefaultWriteCell, comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1], "DEC");
                    }
                }
                else
                {
                    MessageBox.Show("Errore nel settaggio del registro", "Alert");

                    list_coilsTable[(int)(address_start)].Color = Brushes.Red.ToString();
                }
            }
            catch { }
        }

        private void checkBoxViewTableWithoutOffset_CheckedChanged(object sender, RoutedEventArgs e)
        {
            labelOffsetHiddenCoils.Visibility = (bool)checkBoxViewTableWithoutOffset.IsChecked ? Visibility.Visible : Visibility.Hidden;
            labelOffsetHiddenInput.Visibility = (bool)checkBoxViewTableWithoutOffset.IsChecked ? Visibility.Visible : Visibility.Hidden;
            labelOffsetHiddenInputRegister.Visibility = (bool)checkBoxViewTableWithoutOffset.IsChecked ? Visibility.Visible : Visibility.Hidden;
            labelOffsetHiddenHolding.Visibility = (bool)checkBoxViewTableWithoutOffset.IsChecked ? Visibility.Visible : Visibility.Hidden;
        }

        //-----------------------------------------------------------------------------------------
        //------------------FUNZIONE PER AGGIORNARE ELEMENTI GRAFICA DA ALTRI FORM-----------------
        //-----------------------------------------------------------------------------------------

        public void toolSTripMenuEnable(bool enable)
        {
            //simFORMToolStripMenuItem.IsEnabled = enable;
        }

        private void buttonClearAll_Click(object sender, RoutedEventArgs e)
        {
            richTextBoxIncomingPackets.Document.Blocks.Clear();
            richTextBoxIncomingPackets.AppendText("\n");

            richTextBoxOutgoingPackets.Document.Blocks.Clear();
            richTextBoxOutgoingPackets.AppendText("\n");
        }
        
        private void gestisciDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
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

        private void comboBoxHoldingValue06_b_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void menuItemToolBit_Click(object sender, RoutedEventArgs e)
        {
            ComandiBit sim = new ComandiBit(ModBus, this);
            sim.Show();
        }

        private void menuItemToolByte_Click(object sender, RoutedEventArgs e)
        {
            ComandiByte sim_byte = new ComandiByte(ModBus, this);
            sim_byte.Show();
        }

        private void menuItemToolWord_Click(object sender, RoutedEventArgs e)
        {
            ComandiWord sim_word = new ComandiWord(ModBus, this);
            sim_word.Show();
        }

        private void salvaConfigurazioneAttualeNelDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            String prevoiusPath = pathToConfiguration;

            Salva_impianto form_save = new Salva_impianto();
            form_save.ShowDialog();

            //Controllo il risultato del form
            if ((bool)form_save.DialogResult)
            {
                pathToConfiguration = form_save.path;

                salva_configurazione(false);

                if (pathToConfiguration != defaultPathToConfiguration)
                {
                    this.Title = "ModBus C# Client " + version + " - File: " + pathToConfiguration;
                }


                Directory.CreateDirectory("Json\\" + pathToConfiguration);

                String[] fileNames = Directory.GetFiles("Json\\" + prevoiusPath + "\\");

                for (int i = 0; i < fileNames.Length; i++)
                {
                    String newFile = "Json\\" + pathToConfiguration + fileNames[i].Substring(fileNames[i].LastIndexOf('\\'));

                    Console.WriteLine("Copying file: " + fileNames[i] + " to " + newFile);
                    File.Copy(fileNames[i], newFile);
                }
            }
        }

        private void caricaConfigurazioneDaDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Carica_impianto form_load = new Carica_impianto(defaultPathToConfiguration);
            form_load.ShowDialog();

            //Controllo il risultato del form
            if ((bool)form_load.DialogResult)
            {
                salva_configurazione(false);

                pathToConfiguration = form_load.path;

                if (pathToConfiguration != defaultPathToConfiguration)
                {
                    this.Title = "ModBus C# Client " + version + " - File: " + pathToConfiguration;
                }

                carica_configurazione();
            }
        }

        private void apriTemplateEditor_Click(object sender, RoutedEventArgs e)
        {
            TemplateEditor window = new TemplateEditor(this);
            window.ShowDialog();
        }

        private void caricaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            salva_configurazione(false);
            carica_configurazione();
        }

        private void buttonLoopCoils01_Click(object sender, RoutedEventArgs e)
        {
            loopCoils01 = !loopCoils01;
            buttonLoopCoils01.Content = loopCoils01 ? "Stop" : "Loop";
            buttonLoopCoils01.Background = loopCoils01 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }
        private void buttonLoopCoilsRange_Click(object sender, RoutedEventArgs e)
        {
            loopCoilsRange = !loopCoilsRange;
            buttonLoopCoilsRange.Content = loopCoilsRange ? "Stop" : "Loop";
            buttonLoopCoilsRange.Background = loopCoilsRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInput02_Click(object sender, RoutedEventArgs e)
        {
            loopInput02 = !loopInput02;
            buttonLoopInput02.Content = loopInput02 ? "Stop" : "Loop";
            buttonLoopInput02.Background = loopInput02 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInputRange_Click(object sender, RoutedEventArgs e)
        {
            loopInputRange = !loopInputRange;
            buttonLoopInputRange.Content = loopInputRange ? "Stop" : "Loop";
            buttonLoopInputRange.Background = loopInputRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInputRegister04_Click(object sender, RoutedEventArgs e)
        {
            loopInputRegister04 = !loopInputRegister04;
            buttonLoopInputRegister04.Content = loopInputRegister04 ? "Stop" : "Loop";
            buttonLoopInputRegister04.Background = loopInputRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInputRegisterRange_Click(object sender, RoutedEventArgs e)
        {
            loopInputRegisterRange = !loopInputRegisterRange;
            buttonLoopInputRegisterRange.Content = loopInputRegisterRange ? "Stop" : "Loop";
            buttonLoopInputRegisterRange.Background = loopInputRegisterRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopHolding03_Click(object sender, RoutedEventArgs e)
        {
            loopHolding03 = !loopHolding03;
            buttonLoopHolding03.Content = loopHolding03 ? "Stop" : "Loop";
            buttonLoopHolding03.Background = loopHolding03 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopHoldingRange_Click(object sender, RoutedEventArgs e)
        {
            loopHoldingRange = !loopHoldingRange;
            buttonLoopHoldingRange.Content = loopHoldingRange ? "Stop" : "Loop";
            buttonLoopHoldingRange.Background = loopHoldingRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        public void loopPollingRegisters()
        {
            while (loopThreadRunning)
            {
                // Coils
                if (loopCoils01)
                {
                    try
                    {
                        buttonReadCoils01_Click(null, null);
                    }
                    catch(Exception err)
                    {
                        Console.WriteLine(err);
                    }
                }

                if (loopCoilsRange)
                {
                    buttonReadCoilsRange_Click(null, null);
                }

                // Inputs
                if (loopInput02)
                {
                    buttonReadInput02_Click(null, null);
                }

                if (loopInputRange)
                {
                    buttonReadInputRange_Click(null, null);
                }

                // Input Registers
                if (loopInputRegister04)
                {
                    buttonReadInputRegister04_Click(null, null);
                }

                if (loopInputRegisterRange)
                {
                    buttonReadInputRegisterRange_Click(null, null);
                }

                // Holding Registers
                if (loopHolding03)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHolding03_Click(null, null);
                    });
                }

                if (loopHoldingRange)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHoldingRange_Click(null, null);
                    });
                }

                // Pausa tra una query e la successiva
                Thread.Sleep(10000);
            }
        }

        public bool checkLoopStartStop()
        {
            if(loopCoils01 || loopCoilsRange || loopInput02 || loopInputRange || loopInputRegister04 || loopInputRegisterRange || loopHolding03 || loopHoldingRange)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    
        public void checkLoop()
        {
            if (checkLoopStartStop() != loopThreadRunning)
            {
                // Thread già attivo, lo fermo
                if (loopThreadRunning)
                {
                    Console.WriteLine("Fermo il thread di polling");
                    loopThreadRunning = false;

                    Thread.Sleep(100);

                    try
                    {
                        if (threadLoopQuery.IsAlive)
                        {
                            Console.WriteLine("threadLoopQuery Abort");
                            threadLoopQuery.Abort();
                        }
                    }
                    catch (Exception err)
                    {
                        Console.WriteLine(err);
                    }


                }

                // Thread non attivo, lo avvio
                else
                {
                    Console.WriteLine("Avvio il thread di polling");
                    loopThreadRunning = true;
                    threadLoopQuery = new Thread(new ThreadStart(loopPollingRegisters));
                    threadLoopQuery.IsBackground = true;
                    threadLoopQuery.Priority = ThreadPriority.Normal;
                    threadLoopQuery.Start();
                }
            }
        }

        private void TextBoxPollingINterval_TextChanged(object sender, TextChangedEventArgs e)
        {
            int interval = 0;

            if(int.TryParse(TextBoxPollingInterval.Text, out interval))
            {
                if(interval >= 500)
                {
                    pauseLoop = interval;
                }
            }
        }

        private void buttonPingIp_Click(object sender, RoutedEventArgs e)
        {
            Ping p1 = new Ping();
            PingReply PR = p1.Send(textBoxTcpClientIpAddress.Text, 500);

            // check when the ping is not success
            if (!PR.Status.ToString().Equals("Success"))
            {
                buttonPingIp.Background = Brushes.Red;
                DoEvents();
                MessageBox.Show("Ping failed", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                buttonPingIp.Background = Brushes.LightGreen;
                DoEvents();
                MessageBox.Show("Ping ok.\nResponse time: " + PR.RoundtripTime + "ms", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
        }

        private void buttonClearHoldingReg_Click(object sender, RoutedEventArgs e)
        {
            list_holdingRegistersTable.Clear();
            dataGridViewHolding.Items.Refresh();
        }

        private void buttonClearInputReg_Click(object sender, RoutedEventArgs e)
        {
            list_inputRegistersTable.Clear();
            dataGridViewInput.Items.Clear();
        }

        private void buttonClearInput_Click(object sender, RoutedEventArgs e)
        {
            list_inputsTable.Clear();
            dataGridViewInput.Items.Clear();
        }

        private void buttonClearCoils_Click(object sender, RoutedEventArgs e)
        {
            list_coilsTable.Clear();
            dataGridViewCoils.Items.Clear();
        }

        private void licenseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            License window = new License();
            window.Show();
        }

        private void CheckBoxPinWIndow_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = (bool)CheckBoxPinWIndow.IsChecked;
        }

        private void dataGridViewHolding_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (bool)CheckBoxSendValuesOnEditHoldingTable.IsChecked)
            {
                try
                {
                    //dataGridViewHolding.CommitEdit();

                    ModBus_Item currentItem = (ModBus_Item)dataGridViewHolding.SelectedItem;
                    int index = list_holdingRegistersTable.IndexOf(currentItem) - 1;

                    // Se eventualmente fosse da utilizzare il registro precedente usare la seguente
                    currentItem = list_holdingRegistersTable[list_holdingRegistersTable.IndexOf(currentItem) - 1];

                    // Debug
                    Console.WriteLine("Register: " + currentItem.Register);
                    Console.WriteLine("Value: " + currentItem.Value);
                    Console.WriteLine("Notes: " + currentItem.Notes);

                    uint address_start = P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset) + P.uint_parser(currentItem.Register, comboBoxHoldingRegistri);

                    if (address_start > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)    //Se indirizzo espresso in 30001+ imposto offset a 0
                    {
                        address_start = address_start - 40001;
                    }

                    uint value_ = P.uint_parser(currentItem.Value, comboBoxHoldingValori);

                    if (ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress.Text), address_start, value_))
                    {
                        String[] value = { value_.ToString() };

                        list_holdingRegistersTable[index].ValueBin = Convert.ToString(value_ >> 8, 2).PadLeft(8, '0') + " " + Convert.ToString((UInt16)(value_ << 8) >> 8, 2).PadLeft(8, '0'); ;
                        list_holdingRegistersTable[index].Color = colorDefaultWriteCell.ToString();

                        /*if ((bool)checkBoxCreateTableAtBoot.IsChecked)
                        {
                            //Uso le righe esistenti
                            updateRowTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell);
                        }
                        else
                        {
                            //Cancello la tabella e inserisco le nuove righe
                            if ((bool)checkBoxViewTableWithoutOffset.IsChecked)
                                insertRowsTable(list_holdingRegistersTable, null, address_start - P.uint_parser(textBoxHoldingOffset, comboBoxHoldingOffset), value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                            else
                                insertRowsTable(list_holdingRegistersTable, null, address_start, value, colorDefaultWriteCell, comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1], comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1]);
                        }*/
                    }
                    else
                    {
                        MessageBox.Show("Errore nel settaggio del registro", "Alert");

                        list_holdingRegistersTable[(int)(address_start)].Color = Brushes.Red.ToString();
                    }

                    dataGridViewHolding.Items.Refresh();
                    dataGridViewHolding.SelectedIndex = index + 1;
                }
                catch { }
            }
        }

        private void dataGridViewCoils_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (bool)CheckBoxSendValuesOnEditCoillsTable.IsChecked)
            {
                try
                {
                    //dataGridViewHolding.CommitEdit();

                    ModBus_Item currentItem = (ModBus_Item)dataGridViewCoils.SelectedItem;
                    int index = list_coilsTable.IndexOf(currentItem) - 1;

                    // Se eventualmente fosse da utilizzare il registro precedente usare la seguente
                    currentItem = list_coilsTable[list_coilsTable.IndexOf(currentItem) - 1];

                    // Debug
                    Console.WriteLine("Register: " + currentItem.Register);
                    Console.WriteLine("Value: " + currentItem.Value);
                    Console.WriteLine("Notes: " + currentItem.Notes);

                    uint address_start = P.uint_parser(textBoxCoilsOffset, comboBoxCoilsOffset) + P.uint_parser(currentItem.Register, comboBoxCoilsAddress05);

                    if (ModBus.forceSingleCoil_05(byte.Parse(textBoxModbusAddress.Text), address_start, uint.Parse(textBoxCoilsValue05.Text)))
                    {
                        list_coilsTable[index].Color = colorDefaultWriteCell.ToString();
                    }
                    else
                    {
                        MessageBox.Show("Errore nel settaggio del registro", "Alert");

                        list_coilsTable[(int)(address_start)].Color = Brushes.Red.ToString();
                    }
                }
                catch { }
            }
        }
    }

    // Classe per caricare dati dal file di configurazione json
    public class SAVE
    {
        // Variabili interfaccia 
        public bool usingSerial { get; set; } //True -> Serial, False -> TCP

        public string modbusAddress { get; set; }

        //Vatiabili configurazione seriale
        public int serialPort { get; set; }
        public int serialSpeed { get; set; }
        public int serialParity { get; set; }
        public int serialStop { get; set; }

        public bool serialMaster { get; set; } //True -> Master, False -> False
        public bool serialRTU { get; set; }

        //Variabili configurazione tcp
        public string tcpClientIpAddress { get; set; }
        public string tcpClientPort { get; set; }
        public string tcpServerIpAddress { get; set; }
        public string tcpServerPort { get; set; }

        //VARIABILI TABPAGE

        //TabPage1 (Coils)
        public string textBoxCoilsAddress01 { get; set; }
        public string textBoxCoilNumber { get; set; }
        public string textBoxCoilsRange_A { get; set; }
        public string textBoxCoilsRange_B { get; set; }
        public string textBoxCoilsAddress05 { get; set; }
        public string textBoxCoilsValue05 { get; set; }
        public string textBoxCoilsAddress15_A { get; set; }
        public string textBoxCoilsAddress15_B { get; set; }
        public string textBoxCoilsValue15 { get; set; }
        public string textBoxGoToCoilAddress { get; set; }


        //TabPage2 (inputs)
        public string textBoxInputAddress02 { get; set; }
        public string textBoxInputNumber { get; set; }
        public string textBoxInputRange_A { get; set; }
        public string textBoxInputRange_B { get; set; }
        public string textBoxGoToInputAddress { get; set; }

        //TabPage3 (input registers)
        public string textBoxInputRegisterAddress04 { get; set; }
        public string textBoxInputRegisterNumber { get; set; }
        public string textBoxInputRegisterRange_A { get; set; }
        public string textBoxInputRegisterRange_B { get; set; }
        public string textBoxGoToInputRegisterAddress { get; set; }

        //TabPage4 (holding registers)
        public string textBoxHoldingAddress03 { get; set; }
        public string textBoxHoldingRegisterNumber { get; set; }
        public string textBoxHoldingRange_A { get; set; }
        public string textBoxHoldingRange_B { get; set; }
        public string textBoxHoldingAddress06 { get; set; }
        public string textBoxHoldingValue06 { get; set; }
        public string textBoxHoldingAddress16_A { get; set; }
        public string textBoxHoldingAddress16_B { get; set; }
        public string textBoxHoldingValue16 { get; set; }
        public string textBoxGoToHoldingAddress { get; set; }

        //TabPage5 (diagnostic)

        //TabPage6 (summary)
        public bool statoConsole { get; set; }

        //Altri elementi aggiunti dopo
        public string comboBoxCoilsAddress01_ { get; set; }
        public string comboBoxCoilsRange_A_ { get; set; }
        public string comboBoxCoilsRange_B_ { get; set; }
        public string comboBoxCoilsAddress05_ { get; set; }
        public string comboBoxCoilsValue05_ { get; set; }
        public string comboBoxCoilsAddress15_A_ { get; set; }
        public string comboBoxCoilsAddress15_B_ { get; set; }
        public string comboBoxCoilsValue15_ { get; set; }
        public string comboBoxInputAddress02_ { get; set; }
        public string comboBoxInputRange_A_ { get; set; }
        public string comboBoxInputRange_B_ { get; set; }
        public string comboBoxInputRegisterAddress04_ { get; set; }
        public string comboBoxInputRegisterRange_A_ { get; set; }
        public string comboBoxInputRegisterRange_B_ { get; set; }
        public string comboBoxHoldingAddress03_ { get; set; }
        public string comboBoxHoldingRange_A_ { get; set; }
        public string comboBoxHoldingRange_B_ { get; set; }
        public string comboBoxHoldingAddress06_ { get; set; }
        public string comboBoxHoldingValue06_ { get; set; }
        public string comboBoxHoldingAddress16_A_ { get; set; }
        public string comboBoxHoldingAddress16_B_ { get; set; }
        public string comboBoxHoldingValue16_ { get; set; }

        public string comboBoxCoilsRegistri_ { get; set; }
        public string comboBoxInputRegistri_ { get; set; }
        public string comboBoxInputRegRegistri_ { get; set; }
        public string comboBoxInputRegValori_ { get; set; }
        public string comboBoxHoldingRegistri_ { get; set; }
        public string comboBoxHoldingValori_ { get; set; }

        public string comboBoxCoilsAddress05_b_ { get; set; }
        public string comboBoxCoilsValue05_b_ { get; set; }
        public string comboBoxHoldingAddress06_b_ { get; set; }
        public string comboBoxHoldingValue06_b_ { get; set; }

        public string comboBoxCoilsOffset_ { get; set; }
        public string comboBoxInputOffset_ { get; set; }
        public string comboBoxInputRegOffset_ { get; set; }
        public string comboBoxHoldingOffset_ { get; set; }

        public string textBoxCoilsAddress05_b_ { get; set; }
        public string textBoxCoilsValue05_b_ { get; set; }
        public string textBoxHoldingAddress06_b_ { get; set; }
        public string textBoxHoldingValue06_b_ { get; set; }

        public string textBoxCoilsOffset_ { get; set; }
        public string textBoxInputOffset_ { get; set; }
        public string textBoxInputRegOffset_ { get; set; }
        public string textBoxHoldingOffset_ { get; set; }

        public bool checkBoxUseOffsetInTables_ { get; set; }
        public bool checkBoxUseOffsetInTextBox_ { get; set; }
        public bool checkBoxFollowModbusProtocol_ { get; set; }
        public bool checkBoxCreateTableAtBoot_ { get; set; }
        public bool checkBoxSavePackets_ { get; set; }
        public bool checkBoxCloseConsolAfterBoot_ { get; set; }
        public bool checkBoxCellColorMode_ { get; set; }
        public bool checkBoxViewTableWithoutOffset_ { get; set; }

        public string textBoxSaveLogPath_ { get; set; }

        public string comboBoxDiagnosticFunction_ { get; set; }

        public string textBoxDiagnosticFunctionManual_ { get; set; }

        public string colorDefaultReadCell_ { get; set; }
        public string colorDefaultWriteCell_ { get; set; }
        public string colorErrorCell_ { get; set; }

        public string pathToConfiguration_ { get; set; }

        public string TextBoxPollingInterval_ { get; set; }

        public bool? CheckBoxSendValuesOnEditCoillsTable_ { get; set; }
        public bool? CheckBoxSendValuesOnEditHoldingTable_ { get; set; }

    }

    public class dataGridJson
    {
        public string[] registers { get; set; }
        public string[] values { get; set; }
        public string[] note { get; set; }
    }
}
