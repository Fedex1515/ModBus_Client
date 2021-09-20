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

        bool scrolled_log = false;
        int count_log = 0;

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
                            if (count_log > main.LogLimitRichTextBox)
                            {
                                // Arrivato al limite tolgo una riga ogni volta che aggiungo una riga
                                RichTextBoxLog.Document.Blocks.Remove(RichTextBoxLog.Document.Blocks.FirstBlock);
                            }
                            else
                            {
                                count_log += 1;
                            }

                            RichTextBoxLog.AppendText(content);
                        });

                        scrolled_log = false;
                    }
                    else
                    {
                        if (!scrolled_log)
                        {
                            RichTextBoxLog.Dispatcher.Invoke((Action)delegate
                            {
                                RichTextBoxLog.ScrollToEnd();
                            });

                            scrolled_log = true;
                        }

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
            count_log = 0;

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
