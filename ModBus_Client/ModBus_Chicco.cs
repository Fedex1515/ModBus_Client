


// -------------------------------------------------------------------------------------------

// Copyright (c) 2020 Federico Turco

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
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Concurrent;

//Process.
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
using System.IO;

using ModBusMaster_Chicco;

//Classecon funzioni di conversione DEC-HEX
using Raccolta_funzioni_parser;

namespace ModBusMaster_Chicco
{
    public class ModBus_Chicco
    {
        //-----VARIABILI-OGGETTO------
        bool ClientActive = false;

        bool ordineTextBoxLog = true;   // true -> in cima, false -> in fondo

        //RTU o ASCII
        SerialPort serialPort;

        //TCP
        String ip_address;
        String port;

        String type;    //"RTU", "ASCII", "TCP"

        Border pictureBoxSending = new Border();
        Border pictureBoxReceiving = new Border();

        public FixedSizedQueue<String> log = new FixedSizedQueue<string>();
        public FixedSizedQueue<String> log2 = new FixedSizedQueue<string>();

        UInt16 queryCounter = 1;    //Conteggio richieste TCP per inserirli nel primo byte

        int buffer_dimension = 256; //Dimensione buffer per comandi invio/ricezione seriali/tcp

        public int readTimeout = 1000;

        public ModBus_Chicco(SerialPort serialPort_, String ip_address_, String port_, String type_)
        {
            //Type: "TCP", "RTU", "ASCII" (ASCII ANCORA DA IMPLEMENTARE)
            type = type_;

            //RTU/ASCII
            serialPort = serialPort_;

            //TCP
            ip_address = ip_address_;
            port = port_;

            //Dimensione log locale
            log.Limit = 10000;
            log2.Limit = 10000;

            //DEBUG
            Console.WriteLine("Oggeto ModBus:" + type);
        }

        public ModBus_Chicco(SerialPort serialPort_, String ip_address_, String port_, String type_, Border pictureBoxSending_, Border pictureBoxReceiving_)
        {
            //ClientActive = true;
            //Type: TCP, RTU, ASCII
            type = type_;

            //RTU/ASCII
            serialPort = serialPort_;

            //TCP
            ip_address = ip_address_;
            port = port_;

            //GRAFICA
            pictureBoxSending = pictureBoxSending_;
            pictureBoxReceiving = pictureBoxReceiving_;

            //Dimensione log locale
            log.Limit = 10000;
            log2.Limit = 10000;

            //DEBUG
            Console.WriteLine("Oggeto ModBus:" + type);
        }

        public void open()
        {
            ClientActive = true;
        }

        public void close()
        {
            ClientActive = false;
            try
            {
                serialPort.Close();
                //Il client viene chiuso al termine di ogni richiesta
            }
            catch { }

            try
            {

            }
            catch { }
        }

        public String[] readCoilStatus_01(byte slave_add, uint start_add, uint no_of_coils)
        {
            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x01 -> Function
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
             */


            byte[] query;
            byte[] response;
            String[] result = new String[no_of_coils];

            if (type == "TCP" && ClientActive)
            {
                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x01;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_coils >> 8);
                query[11] = (byte)(no_of_coils);

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //Metto dopo la pictureBox rispetto a stream.read perche' se si pianta nella
                //letttura la picturebox rimarrebbe gialla

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //Leggo i bit di ciascun byte partendo dal 9 che contiene le prime 8 coils
                //La coil 0 e nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                for (int i = 9; i < Length; i += 1)
                {
                    for (int a = 0; a < 8; a++)
                    {
                        try
                        {
                            //Se supero l'indice me ne frego (accade se coil % 8 != 0)
                            //che tanto va nel catch

                            //DEBUG
                            Console.WriteLine("i: " + i.ToString() + " a: " + a.ToString());

                            result[(i - 9) * 8 + a] = Convert.ToInt32((response[i] & (1 << a)) > 0).ToString();
                        }
                        catch
                        {
                            //result[(i - 9) * 8 + a] = "?";
                        }
                    }
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console.WriteLine("Result (array of coils): " + result);

                return result;

            }
            else if (type == "RTU" && ClientActive)
            {
                /*
                RTU
                0x07 -> Slave Address
                0x01 -> Function
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x01;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_coils >> 8);
                query[5] = (byte)(no_of_coils);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                //Pausa per aspettare che arrivi la risposta sul buffer
                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                response = new Byte[buffer_dimension];

                int Length = 0;

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                    Console_printByte("Received: ", response, Length);
                    Console_print(" rx <- ", response, Length);

                    //Leggo i bit di ciascun byte partendo dal 3 che contiene le prime 8 coils
                    //La coil 0 e' nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                    for (int i = 3; i < Length; i += 1)
                    {
                        for (int a = 0; a < 8; a++)
                        {
                            try
                            {
                                //Se supero l'indice me ne frego (accade se coil % 8 != 0)
                                //che tanto va nel catch

                                //DEBUG
                                Console.WriteLine("i: " + i.ToString() + " a: " + a.ToString());

                                result[(i - 3) * 8 + a] = Convert.ToInt32((response[i] & (1 << a)) > 0).ToString();

                                //DEBUG
                                //Console.WriteLine(((response[i] & (1 << a)) > 0).ToString());
                                //Console.WriteLine(response[i].ToString());
                                //Console.WriteLine((1 << a).ToString());
                            }
                            catch
                            {
                                result[(i - 3) * 8 + a] = "?";
                            }
                        }
                    }

                    Console.WriteLine("Result (array of coils): " + result);
                }
                catch
                {
                    Console.WriteLine("Result (array of coils): " + result);
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (!Check_CRC(response, Length))
                {
                    MessageBox.Show("Errore crc pacchetto ricevuto", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                return result;

            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return null;
            }
        }

        public String[] readInputStatus_02(byte slave_add, uint start_add, uint no_of_input)
        {
            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x02 -> Function Code
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
             */

            byte[] query;
            byte[] response;
            String[] result = new String[no_of_input];

            if (type == "TCP" && ClientActive)
            {
                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x02;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_input >> 8);
                query[11] = (byte)(no_of_input);

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //Leggo i bit di ciascun byte partendo dal 9 che contiene le prime 8 coils
                //La coil 0 e nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                for (int i = 9; i < Length; i += 1)
                {
                    for (int a = 0; a < 8; a++)
                    {
                        try
                        {
                            //Se supero l'indice del max me ne frego (accade se coil % 8 != 0)
                            //che tanto va nel catch

                            //DEBUG
                            Console.WriteLine("i: " + i.ToString() + " a: " + a.ToString());

                            result[(i - 9) * 8 + a] = Convert.ToInt32((response[i] & (1 << a)) > 0).ToString();
                        }
                        catch
                        {
                            ;
                        }
                    }
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console.WriteLine("Result (array of inputs): " + result);

                return result;

            }
            else if (type == "RTU" && ClientActive)
            {
                /*
                RTU
                0x07 -> Slave Address
                0x02 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x02;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_input >> 8);
                query[5] = (byte)(no_of_input);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                    Console_printByte("Received: ", response, Length);
                    Console_print(" rx <- ", response, Length);

                    //Leggo i bit di ciascun byte partendo dal 9 che contiene le prime 8 coils
                    //La coil 0 e nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                    for (int i = 3; i < Length; i += 1)
                    {
                        for (int a = 0; a < 8; a++)
                        {
                            try
                            {
                                //Se supero l'indice me ne frego (accade se coil % 8 != 0)
                                //che tanto va nel catch
                                result[(i - 3) * 8 + a] = Convert.ToInt32((response[i] & (1 << a)) > 0).ToString();

                                //DEBUG
                                //Console.WriteLine(((response[i] & (1 << a)) > 0).ToString());
                                //Console.WriteLine(response[i].ToString());
                                //Console.WriteLine((1 << a).ToString());
                            }
                            catch
                            {
                                ;
                            }
                        }
                    }

                    Console.WriteLine("Result (array of inputs): " + result);
                }
                catch
                {
                    Console.WriteLine("Result (array of inputs): " + result);
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (!Check_CRC(response, Length))
                {
                    MessageBox.Show("Errore crc pacchetto ricevuto", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                return result;
            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return null;
            }
        }

        public String[] readHoldingRegister_03(byte slave_add, uint start_add, uint no_of_registers)
        {
            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x03 -> Function Code
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
             */

            byte[] query;
            byte[] response;
            String[] result = new String[no_of_registers];

            if (type == "TCP" && ClientActive)
            {
                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x03;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_registers >> 8);
                query[11] = (byte)(no_of_registers);

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                //Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                //------------------------------------------

                //------------pictureBox gialla-------------
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //Thread.Sleep(50);

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                for (int i = 9; i < Length; i += 2)
                {
                    result[(i - 9) / 2] = ((uint)(response[i] << 8) + (uint)(response[i + 1])).ToString();
                }

                //------------pictureBox grigia-------------
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //Thread.Sleep(50);
                //------------------------------------------

                Console.WriteLine("Result (array of registers): " + result);

                return result;

            }
            else if (type == "RTU" && ClientActive)
            {
                /*
                RTU
                0x07 -> Slave Address
                0x03 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x03;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_registers >> 8);
                query[5] = (byte)(no_of_registers);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                    Console_printByte("Received: ", response, Length);
                    Console_print(" rx <- ", response, Length);

                    for (int i = 3; i < Length - 2; i += 2)
                    {
                        result[(i - 3) / 2] = ((uint)(response[i] << 8) + (uint)(response[i + 1])).ToString();
                    }

                    Console.WriteLine("Result (array of registers): " + result);


                }
                catch
                {
                    Console.WriteLine("Errore lettura porta seriale");
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (!Check_CRC(response, Length))
                {
                    MessageBox.Show("Errore crc pacchetto ricevuto", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                return result;
            }
            else
            {
                string[] error = { "?" };
                Console.WriteLine("Nessuna connessione attiva");
                return error;
            }
        }

        public String[] readInputRegister_04(byte slave_add, uint start_add, uint no_of_registers)
        {

            byte[] query;
            byte[] response;
            String[] result = new String[no_of_registers];

            if (type == "TCP" && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x04;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_registers >> 8);
                query[11] = (byte)(no_of_registers);

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------


                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------


                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                for (int i = 9; i < Length; i += 2)
                {
                    result[(i - 9) / 2] = ((uint)(response[i] << 8) + (uint)(response[i + 1])).ToString();
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console.WriteLine("Result (array of registers): " + result);

                return result;

            }
            else if (type == "RTU" && ClientActive)
            {
                /*
                RTU:
                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x04;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_registers >> 8);
                query[5] = (byte)(no_of_registers);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                    Console_printByte("Received: ", response, Length);
                    Console_print(" rx <- ", response, Length);

                    for (int i = 3; i < Length - 2; i += 2) //-2 di CRC
                    {
                        result[(i - 3) / 2] = ((uint)(response[i] << 8) + (uint)(response[i + 1])).ToString();
                    }

                    Console.WriteLine("Result (array of registers): " + result);


                }
                catch
                {
                    Console.WriteLine("Timeout lettura porta seriale");

                    for (int i = 0; i < result.Length; i += 1)
                    {
                        result[i] = "?";
                    }
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (!Check_CRC(response, Length))
                {
                    MessageBox.Show("Errore crc pacchetto ricevuto", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                return result;
            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return null;
            }
        }

        public bool forceSingleCoil_05(byte slave_add, uint start_add, uint state)
        {
            //True se la funzione riceve risposta affermativa

            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x05 -> Function Code
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi (0xFF -> On, 0x00 -> Off)
            0x03 -> No of Registers Lo (Sempre 0x00)
             */

            byte[] query;
            byte[] response;
            //uint[] result = new uint[no_of_registers];

            if (type == "TCP" && ClientActive)
            {
                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;    //Message Length
                query[5] = 0x06;    //Message Length

                query[6] = slave_add;
                query[7] = 0x05;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                if (state > 0)
                    query[10] = 0xFF;
                else
                    query[10] = 0x00;

                query[11] = 0x00;

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (Length == query.Length)
                    return true;

                return false;

            }
            else if (type == "RTU" && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x05 -> Header
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi (0xFF -> On, 0x00 -> Off)
                0x03 -> No of Registers Lo (Sempre 0x00)
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                 */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x05;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                if (state > 0)
                    query[4] = 0xFF;
                else
                    query[4] = 0x00;

                query[5] = 0x00;

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                response = new Byte[buffer_dimension];
                int Length = new int();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);
                }
                catch
                {
                    return false;
                }

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                return Length == query.Length && Check_CRC(response, Length);

            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return false;
            }
        }

        public bool presetSingleRegister_06(byte slave_add, uint start_add, uint value)
        {
            //True se la funzione riceve risposta affermativa

            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x06 -> FUnction Code
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
             */

            byte[] query;
            byte[] response;

            if (type == "TCP" && ClientActive)
            {
                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x06;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                query[10] = (byte)(value >> 8);
                query[11] = (byte)(value);

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (Length == query.Length)
                    return true;

                return false;

            }
            else if (type == "RTU" && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x06 -> Header
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x06;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                query[4] = (byte)(value >> 8);
                query[5] = (byte)(value);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                }
                catch
                {
                    return false;
                }

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                return Length == query.Length && Check_CRC(response, Length);
            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return false;
            }
        }

        public String diagnostics_08(byte slave_add, uint sub_func, uint data)
        {

            byte[] query;
            byte[] response;
            String[] result = new String[buffer_dimension];

            if (type == "TCP" && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Subfunction Hi
                0x2C -> Subfunction LO
                0x00 -> Data Hi
                0x03 -> Data Lo
                */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x08;
                query[8] = (byte)(sub_func >> 8);
                query[9] = (byte)(sub_func);
                query[10] = (byte)(data >> 8);
                query[11] = (byte)(data);

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------


                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------


                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console.WriteLine("Result (array of registers): " + result);

                string result_ = "";
                result = new String[Length];

                for (int i = 0; i < Length; i++)
                {
                    try
                    {
                        //if(result[i] < 10)
                        result[i] = response[i].ToString("X");

                        if (int.Parse(result[i]) < 10)
                            result_ += "0";

                        result_ += response[i].ToString("X") + " ";
                    }
                    catch
                    {
                        Console.WriteLine("Errore analisi response diagnostic function");
                    }
                }

                return result_;

            }
            else if (type == "RTU" && ClientActive)
            {
                /*
                RTU:
                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Subfunction Hi
                0x2C -> Subfunction Lo
                0x00 -> Data Hi
                0x03 -> Data Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x08;
                query[2] = (byte)(sub_func >> 8);
                query[3] = (byte)(sub_func);
                query[4] = (byte)(data >> 8);
                query[5] = (byte)(data);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                    Console_printByte("Received: ", response, Length);
                    Console_print(" rx <- ", response, Length);

                    Console.WriteLine("Result (array of registers): " + result);


                }
                catch
                {
                    Console.WriteLine("Timeout lettura porta seriale");

                }

                if(!Check_CRC(response, Length))
                {
                    MessageBox.Show("Errore crc pacchetto ricevuto", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                string result_ = "";
                result = new String[Length];

                for (int i = 0; i < Length; i++)
                {
                    try
                    {
                        //if(result[i] < 10)
                        result[i] = response[i].ToString("X");

                        if (int.Parse(result[i]) < 10)
                            result_ += "0";

                        result_ += response[i].ToString("X") + " ";
                    }
                    catch
                    {
                        Console.WriteLine("Errore analisi response diagnostic function");
                    }
                }

                return result_;
            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return null;
            }
        }

        public bool forceMultipleCoils_15(byte slave_add, uint start_add, bool[] coils_value)
        {
            // True se la funzione riceve risposta affermativa

            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x15 -> Function Code
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
            
            0x02 -> Byte count

            0x00 -> Data Hi
            0x00 -> Data Lo
             */

            byte[] query;
            byte[] response;

            if (type == "TCP" && ClientActive)
            {
                queryCounter++;
                query = new byte[13 + (coils_value.Length/8) + (coils_value.Length % 2 == 0 ? 0 : 1)];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x0F;

                // Starting address
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                // Number of regsiters
                query[10] = (byte)(coils_value.Length >> 8);
                query[11] = (byte)(coils_value.Length);

                // Byte count
                query[12] = (byte)((coils_value.Length / 8) + (coils_value.Length % 2 == 0 ? 0 : 1));

                for (int i = 0; i < (coils_value.Length / 8 + coils_value.Length % 2 == 0 ? 0 : 1) ; i++)
                {
                    byte val = 0;

                    for(int a = 0; a < 8; a++)
                    {
                        if (a + i * 8 < coils_value.Length)
                        {
                            if(coils_value[a + i * 8])
                            {
                                val += (byte)(1 << a);

                                // debug
                                Console.WriteLine("coil " + a.ToString() + ":1");
                            }
                            else
                            {
                                // debug
                                Console.WriteLine("coil " + a.ToString() + ":0");
                            }
                        }
                    }

                    query[13 + i] = val;

                    // debug
                    Console.WriteLine("byte: " +  val.ToString());
                }

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (response.Length > 11)
                {
                    if (response[8] == query[8] &&
                        response[9] == query[9] &&
                        response[10] == query[10] &&
                        response[11] == query[11])
                    {
                        return true;
                    }
                }

                return false;

            }
            else if (type == "RTU" && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x06 -> Header
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x02 -> Byte count
                0x00 -> Data[0] Hi
                0x03 -> Data[0] Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[6 + coils_value.Length * 2];

                query[0] = slave_add;
                query[1] = 0x10;

                // Starting address
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                // Number of regsiters
                query[4] = (byte)(coils_value.Length >> 8);
                query[5] = (byte)(coils_value.Length);

                // Byte count
                query[6] = (byte)(coils_value.Length * 2);

                for (int i = 0; i < coils_value.Length; i++)
                {
                    //query[7 + 2 * i] = (byte)(coils_value[i] >> 8);
                    //query[8 + 2 * i] = (byte)(coils_value[i]);
                }

                byte[] crc = Calcolo_CRC(query, 7 + coils_value.Length * 2);

                query[7 + coils_value.Length * 2] = crc[0];
                query[8 + coils_value.Length * 2] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                }
                catch
                {
                    return false;
                }

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                return Check_CRC(response, Length);
            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return false;
            }
        }

        public bool presetMultipleRegisters_16(byte slave_add, uint start_add, uint[] register_value)
        {
            // True se la funzione riceve risposta affermativa

            /*
            TCP:
            0x00
            0x01
            0x00
            0x00
            0x00 -> Message Length Hi
            0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

            0x07 -> Slave Address
            0x16 -> Function Code
            0x01 -> Start Addr Hi
            0x2C -> Start Addr Lo
            0x00 -> No of Registers Hi
            0x03 -> No of Registers Lo
            
            0x02 -> Byte count

            0x00 -> Data Hi
            0x00 -> Data Lo
             */

            byte[] query;
            byte[] response;

            if (type == "TCP" && ClientActive)
            {
                queryCounter++;
                query = new byte[13 + register_value.Length*2];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x10;

                // Starting address
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                // Number of regsiters
                query[10] = (byte)(register_value.Length >> 8);
                query[11] = (byte)(register_value.Length);

                // Byte count
                query[12] = (byte)(register_value.Length * 2);

                for (int i = 0; i < register_value.Length; i++)
                {
                    query[13 + 2*i] = (byte)(register_value[i] >> 8);
                    query[14 + 2*i] = (byte)(register_value[i]);
                }

                TcpClient client = new TcpClient(ip_address, int.Parse(port));

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                NetworkStream stream = client.GetStream();
                stream.ReadTimeout = readTimeout;
                stream.Write(query, 0, query.Length);

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                response = new Byte[buffer_dimension];
                int Length = stream.Read(response, 0, response.Length);
                client.Close();

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                if (response.Length > 11) {
                    if (response[8] == query[8] &&
                        response[9] == query[9] &&
                        response[10] == query[10] &&
                        response[11] == query[11])
                    {
                        return true;
                    }
                }

                return false;

            }
            else if (type == "RTU" && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x06 -> Header
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x02 -> Byte count
                0x00 -> Data[0] Hi
                0x03 -> Data[0] Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[9 + register_value.Length*2];

                query[0] = slave_add;
                query[1] = 0x10;

                // Starting address
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                // Number of regsiters
                query[4] = (byte)(register_value.Length >> 8);
                query[5] = (byte)(register_value.Length);

                // Byte count
                query[6] = (byte)(register_value.Length * 2);

                for (int i = 0; i < register_value.Length; i++)
                {
                    query[7 + 2 * i] = (byte)(register_value[i] >> 8);
                    query[8 + 2 * i] = (byte)(register_value[i]);
                }

                byte[] crc = Calcolo_CRC(query, 7 + register_value.Length*2);

                query[7 + register_value.Length * 2] = crc[0];
                query[8 + register_value.Length * 2] = crc[1];

                //------------pictureBox gialla-------------
                pictureBoxSending.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                //------------pictureBox grigia-------------
                pictureBoxSending.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                Console_printByte("Sent: ", query, query.Length);
                Console_print(" tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                //------------pictureBox gialla-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.Yellow;
                DoEvents();
                //------------------------------------------

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);

                }
                catch
                {
                    return false;
                }

                Console_printByte("Received: ", response, Length);
                Console_print(" rx <- ", response, Length);

                //------------pictureBox grigia-------------
                Thread.Sleep(50);
                pictureBoxReceiving.Background = Brushes.LightGray;
                DoEvents();
                //------------------------------------------

                return Check_CRC(response, Length);
            }
            else
            {
                Console.WriteLine("Nessuna connessione attiva");
                return false;
            }
        }

        //-----------------------------------------------------------------
        //--------------------Calcolo CRC 16 MODBUS------------------------
        //-----------------------------------------------------------------

        // Calcolo CRC MODBUS
        public byte[] Calcolo_CRC(byte[] message, int length)
        {
            UInt16 crc = 0xFFFF;
            byte[] result = new byte[2];

            for (int pos = 0; pos < length; pos++)
            {
                crc ^= (UInt16)message[pos];    //XOR

                for (int i = 8; i != 0; i--)
                {
                    // Passo ogni byte del pacchetto
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {             
                        crc >>= 1;
                    }
                }
            }

            result[0] = (byte)(crc);        //LSB
            result[1] = (byte)(crc >> 8);   //MSB

            return result;
        }
        
        bool Check_CRC(byte[] message, int length)
        {
            UInt16 crc = 0xFFFF;
            byte[] result = new byte[2];

            for (int pos = 0; pos < (length - 2); pos++)
            {
                crc ^= (UInt16)message[pos];    //XOR

                for (int i = 8; i != 0; i--)
                {
                    // Passo ogni byte del pacchetto
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {             
                        crc >>= 1;
                    }
                }
            }

            result[0] = (byte)(crc);        //LSB
            result[1] = (byte)(crc >> 8);   //MSB

            return ((byte)(crc) == message[length - 2] && ((byte)(crc >> 8) == message[length - 1]));
        }


        public string timestamp()
        {
            return DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" +
                   DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" +
                   DateTime.Now.Second.ToString().PadLeft(2, '0');
        }

        // Cambia ordine inserimento righe nella RichTextBox di logi
        public void insertLogLinesAtTop(bool yes)
        {
            ordineTextBoxLog = yes;
        }

        //-------------------------------------------------------------------------------------
        //----------------Funzioni stampa su console o textBox array di byte-------------------
        //------------------------------------------------------------------------------------
        private void RichTextBox_printByte(RichTextBox textBox, byte[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {

                    aa = query[i].ToString("X");

                    if (aa.Length < 2)
                        aa = "0" + aa;

                    message += "0x" + aa + " ";
                }

                textBox.AppendText(message + "\n");
            }

        }

        private void Console_printByte(String intestazione, byte[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {

                    aa = query[i].ToString("X");

                    if (aa.Length < 2)
                        aa = "0" + aa;

                    message += "0x" + aa + " ";
                }
                Console.WriteLine(intestazione + message);
            }
        }

        public string Console_print(string header, byte[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {

                    aa = query[i].ToString("X");

                    if (aa.Length < 2)
                        aa = "0" + aa;

                    message += "" + aa + " ";
                }

                log.Enqueue(timestamp() + header + message + "\n");
                log2.Enqueue(timestamp() + header + message + "\n");

                return timestamp() + header + message + "\n";
            }
            else
            {
                return "";
            }
        }

        private void Console_printUint(String intestazione, uint[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {
                    aa = query[i].ToString("X");

                    if (aa.Length < 2)
                        aa = "0" + aa;

                    message += "0x" + aa + " ";
                }
                Console.WriteLine(intestazione + message);
            }
        }

        private void Console_printBool(String intestazione, bool[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {
                    if (query[i] == true)
                        aa = "1";
                    else
                        aa = "0";

                    message += "" + aa + " ";
                }
                Console.WriteLine(intestazione + message);
            }
        }

        // Funzione equivalnete alla vecchia Application.DoEvents()
        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
        }
    }

    public class ModBus_Item
    {
        public string Register { get; set; }
        public string Value { get; set; }
        public string ValueBin { get; set; }
        public string Notes { get; set; }
        public string Color { get; set; }
    }

    public class FixedSizedQueue<T>
    {
        ConcurrentQueue<T> q = new ConcurrentQueue<T>();
        private object lockObject = new object();

        public int Limit { get; set; }
        public void Enqueue(T obj)
        {
            q.Enqueue(obj);

            lock (lockObject)
            {
                T overflow;
                while (q.Count > Limit && q.TryDequeue(out overflow)) ;
            }
        }

        public bool TryDequeue(out T obj)
        {
            if(q.TryDequeue(out obj))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
