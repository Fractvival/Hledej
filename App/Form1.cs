using System;
using System.IO;
using System.IO.Ports;
using System.Windows.Forms;


namespace Hledej
{


    public partial class Form1 : Form
    {

        SerialPort port;
        ComSetting serialSetting;

        public Form1()
        {
            InitializeComponent();
            this.Load += MainForm_Load;
            this.Shown += MainForm_Show;
            this.FormClosing += MainForm_FormClosing;
        }

        class ComSetting
        {
            public string PortName { get; set; }
            public int BaudRate { get; set; }
            public int DataBits { get; set; }
            public Parity Parity { get; set; }
            public StopBits StopBits { get; set; }
        }


        ComSetting ReadComSetting(string filePath)
        {
            ComSetting comSetting = new ComSetting();

            try
            {
                using (StreamReader sr = new StreamReader(filePath))
                {
                    int lineNumber = 1;
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (lineNumber % 2 == 0)
                        {
                            switch (lineNumber)
                            {
                                case 2:
                                    comSetting.PortName = line;
                                    break;
                                case 4:
                                    comSetting.BaudRate = int.Parse(line);
                                    break;
                                case 6:
                                    comSetting.DataBits = int.Parse(line);
                                    break;
                                case 8:
                                {
                                    if (line=="0")
                                    {
                                        comSetting.Parity = Parity.None;
                                    }
                                    else if (line=="1")
                                    {
                                        comSetting.Parity = Parity.Odd;
                                    }
                                    else if (line=="2")
                                    {
                                        comSetting.Parity = Parity.Even;
                                    }
                                    else if (line=="3")
                                    {
                                        comSetting.Parity = Parity.Mark;
                                    }
                                    else if (line == "4")
                                    {
                                        comSetting.Parity = Parity.Space;
                                    }
                                    break;
                                }
                                case 10:
                                {
                                    if (line == "0")
                                    {
                                        comSetting.StopBits = StopBits.None;
                                    }
                                    else if (line == "1")
                                    {
                                        comSetting.StopBits = StopBits.One;
                                    }
                                    else if (line == "2")
                                    {
                                        comSetting.StopBits = StopBits.Two;
                                    }
                                    else if (line == "3")
                                    {
                                        comSetting.StopBits = StopBits.OnePointFive;
                                    }
                                    break;
                                }
                            }
                        }
                        lineNumber++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chyba při čtení souboru: {ex.Message}");
            }

            return comSetting;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            serialSetting = ReadComSetting(@"comset.txt");
            try
            {
                port = new SerialPort(
                    serialSetting.PortName,
                    serialSetting.BaudRate,
                    serialSetting.Parity,
                    serialSetting.DataBits,
                    serialSetting.StopBits);
                port.Open();
                if (!port.IsOpen)
                {
                    MessageBox.Show(@"Port COM nejde otevrit!", @"Chyba na COM portu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Problem se souborem comset.txt pro nastaveni COM portu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void MainForm_Show(object sender, EventArgs e)
        {
        }

        private void ComReceived(object sender, SerialDataReceivedEventArgs e)
        {

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (port.IsOpen)
                {
                    port.Close();
                    // Počkejte, až bude port skutečně uzavřen
                    while (port.IsOpen)
                    {
                        // Počkejte krátkou dobu, než se znovu zkontroluje stav portu
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Problem s uzavrenim COM portu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


    }
}
