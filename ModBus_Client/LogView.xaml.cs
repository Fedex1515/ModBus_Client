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

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for LogView.xaml
    /// </summary>
    public partial class LogView : Window
    {
        MainWindow main;
        bool exit = false;

        Thread threadDequeue;

        public LogView(MainWindow main_)
        {
            main = main_;

            InitializeComponent();

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth) - (windowWidth) - 10;
            this.Top = (screenHeight) - (windowHeight) - 10;
        }

        public void Dequeue()
        {
            while (!exit)
            {
                String content;

                if (main.ModBus != null)
                {
                    if (main.ModBus.log.TryDequeue(out content))
                    {
                        RichTextBoxLog.Dispatcher.Invoke((Action)delegate
                        {
                            RichTextBoxLog.AppendText(content);
                            RichTextBoxLog.ScrollToEnd();
                        });
                    }
                    else
                    {
                        Thread.Sleep(100);
                    }
                }
                else
                {
                    Thread.Sleep(100);
                }
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            RichTextBoxLog.AppendText("\n");
            RichTextBoxLog.Document.PageWidth = 5000;

            // Metto la finestra in primo piano
            CheckBoxPinWindowLog.IsChecked = true;
            this.Topmost = (bool)CheckBoxPinWindowLog.IsChecked;

            threadDequeue = new Thread(new ThreadStart(Dequeue));
            threadDequeue.IsBackground = true;
            threadDequeue.Start();

            main.Focus();
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete) {
                RichTextBoxLog.Document.Blocks.Clear();
                RichTextBoxLog.AppendText("\n");
            }
        }

        private void ButtonClearLog_Click(object sender, RoutedEventArgs e)
        {
            RichTextBoxLog.Document.Blocks.Clear();
            RichTextBoxLog.AppendText("\n");
        }

        private void CheckBoxPinWindowLog_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = (bool)CheckBoxPinWindowLog.IsChecked;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            main.logWindowIsOpen = false;
            exit = true;
        }
    }
}
