using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.IO.Ports;
using System.Windows;
using System.Threading;
using System.Drawing;
using System.Windows.Controls;

namespace Raccolta_funzioni_parser
{
    class Parser    // Classe di conversioni personalizzate
    {
        public uint uint_parser(TextBox textBox, ComboBox comboBox)
        {
            // debug
            // Console.WriteLine(comboBox.SelectedValue.ToString().Split(' ')[1]);

            if (comboBox.SelectedValue.ToString().Split(' ')[1] == "HEX")
            {
                //Numero passato in hex
                try
                {
                    return UInt16.Parse(textBox.Text, System.Globalization.NumberStyles.HexNumber);
                }
                catch
                {
                    //MessageBox.Show("Valore inserito non valido", "Alert");
                    return 0;
                }
            }
            else
            {
                //Numero passato in decimale
                try
                {
                    return UInt16.Parse(textBox.Text);
                }
                catch
                {
                    //MessageBox.Show("Valore inserito non valido", "Alert");
                    return 0;
                }
            }
        }

        public uint uint_parser(String textBox, ComboBox comboBox)
        {
            // debug
            // Console.WriteLine(comboBox.SelectedValue.ToString().Split(' ')[1]);

            if (comboBox.SelectedValue.ToString().Split(' ')[1] == "HEX")
            {
                //Numero passato in hex
                try
                {
                    return UInt16.Parse(textBox, System.Globalization.NumberStyles.HexNumber);
                }
                catch
                {
                    //MessageBox.Show("Valore inserito non valido", "Alert");
                    return 0;
                }
            }
            else
            {
                //Numero passato in decimale
                try
                {
                    return UInt16.Parse(textBox);
                }
                catch
                {
                    //MessageBox.Show("Valore inserito non valido", "Alert");
                    return 0;
                }
            }
        }
    }
}
