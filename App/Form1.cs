﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;


namespace Hledej
{


    public partial class Form1 : Form
    {

        static SerialPort port;
        static ComSetting serialSetting;
        static List<Part> parts;

        public Form1()
        {
            InitializeComponent();
            this.Load += MainForm_Load;
            this.Shown += MainForm_Show;
            this.FormClosing += MainForm_FormClosing;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Nastavení LicenseContext
        }

        class ComSetting
        {
            public string PortName { get; set; }
            public int BaudRate { get; set; }
            public int DataBits { get; set; }
            public Parity Parity { get; set; }
            public StopBits StopBits { get; set; }
        }

        class Part
        {
            public string KZM { get; set; }
            public string PartNumber { get; set; }
            public string Nazev { get; set; }
            public string Pocet { get; set; }
            public string PocetInventura { get; set; }
            public string Umisteni { get; set; }
            public string Doplneno { get; set; }
        }


        static List<Part> ReadPartsFromExcel(string filePath)
        {
            List<Part> parts = new List<Part>();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // První list
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                for (int row = 2; row <= rowCount; row++) // Začínáme na druhém řádku, předpokládáme záhlaví na prvním řádku
                {
                    Part part = new Part();
                    part.KZM = worksheet.Cells[row, 1].Value?.ToString();
                    part.PartNumber = worksheet.Cells[row, 2].Value?.ToString();
                    part.Nazev = $"{worksheet.Cells[row, 3].Value?.ToString()} {worksheet.Cells[row, 4].Value?.ToString()}".Trim(); // Sloučení Název1 a Název2 s mezerou
                    part.Pocet = worksheet.Cells[row, 5].Value?.ToString();
                    part.PocetInventura = worksheet.Cells[row, 6].Value?.ToString();
                    part.Umisteni = worksheet.Cells[row, 7].Value?.ToString();
                    part.Doplneno = worksheet.Cells[row, 8].Value?.ToString();
                    // Přidání části do seznamu
                    parts.Add(part);
                }
            }
            return parts;
        }


        static Part FindPart(string partNumberOrKZM)
        {
            Part foundPart = parts.Find(p => p.PartNumber == partNumberOrKZM);

            if (foundPart == null)
            {
                foundPart = parts.Find(p => p.KZM == partNumberOrKZM);
            }

            if (foundPart == null)
            {
                foundPart = new Part { KZM = "0" }; // Pokud se nic nenajde, vrátí Part s KZM = 0
            }

            return foundPart;
        }


        static List<Part> FindPartsByName(string Name)
        {
            List<Part> foundParts = parts.FindAll(p => p.Nazev.Contains(Name));
            return foundParts;
        }

        static List<Part> FindPartsByAllName(string Name)
        {
            List<Part> foundParts = new List<Part>();

            foreach (var part in parts)
            {
                if (part.Nazev.Equals(Name, StringComparison.OrdinalIgnoreCase) ||
                    part.Nazev.IndexOf(Name, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundParts.Add(part);
                }
            }

            return foundParts;
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
            parts = ReadPartsFromExcel(@"Sklad.xlsx");
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

        private void find_Click(object sender, EventArgs e)
        {
            if (findText.TextLength == 0)
            {
                Console.Beep(233, 80);
                Console.Beep(233, 80);
            }
            else
            {
                Part tryFind = FindPart(findText.Text);
                if (tryFind.KZM == "0")
                {
                    List<Part> listParts = FindPartsByAllName(findText.Text);
                    if (listParts.Count == 0)
                    {
                        Console.Beep(233, 80);
                        Console.Beep(233, 80);
                    }
                    else
                    {
                        listBox1.Items.Clear();
                        for (int i = 0; i < listParts.Count; i++)
                        {
                            listBox1.Items.Add("KZM: " + listParts[i].KZM + " | " + 
                                "PN: " + listParts[i].PartNumber + " | " +
                                "Nazev: " + listParts[i].Nazev + " | " +
                                "Misto: " + listParts[i].Umisteni);
                        }
                        listBox1.Show();
                        buttoncloselist.Show();
                    }
                }
                else 
                { 

                }
            }
        }

        private void buttoncloselist_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.Hide();
            buttoncloselist.Hide();
        }
    }
}
