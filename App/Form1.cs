using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


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
                try
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
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, @"Problem se souborem skladu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Environment.Exit(0);
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
                    MessageBox.Show(@"Port COM nejde otevřít!", @"Chyba na COM portu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    port.DataReceived += new SerialDataReceivedEventHandler(ComReceived);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Problém s nastavením COM portu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            parts = ReadPartsFromExcel(@"Sklad.xlsx");
        }

        private void MainForm_Show(object sender, EventArgs e)
        {
        }


        public static class TextBoxHelper
        {
            public static void MoveCursorToEnd(System.Windows.Forms.TextBox textBox)
            {
                if (textBox == null) return;
                textBox.BeginInvoke((MethodInvoker)delegate {
                    textBox.SelectionStart = textBox.Text.Length;
                    textBox.ScrollToCaret();
                });
            }

            public static void AppendTextAtCursorAndMoveCursorToEnd(System.Windows.Forms.TextBox textBox, string text)
            {
                if (textBox == null || text == null) return;

                if (textBox.InvokeRequired)
                {
                    textBox.Invoke(new Action(() => AppendTextAtCursorAndMoveCursorToEnd(textBox, text)));
                }
                else
                {
                    // Uložení aktuálního pozice kurzoru
                    int cursorPosition = textBox.SelectionStart;

                    // Rozdělení textu na část před kurzorem a část za kurzorem
                    string textBeforeCursor = textBox.Text.Substring(0, cursorPosition);
                    string textAfterCursor = textBox.Text.Substring(cursorPosition);

                    // Vložení nového textu na aktuální pozici kurzoru
                    textBox.Text = textBeforeCursor + text + textAfterCursor;

                    // Nastavení kurzoru za přidaný text
                    textBox.SelectionStart = cursorPosition + text.Length;
                    textBox.ScrollToCaret();
                }
            }


            public static void RemoveCharacterAtCursor(System.Windows.Forms.TextBox textBox)
            {
                if (textBox == null || textBox.Text.Length == 0) return;

                if (textBox.InvokeRequired)
                {
                    textBox.Invoke(new Action(() => RemoveCharacterAtCursor(textBox)));
                }
                else
                {
                    // Uložení aktuální pozice kurzoru
                    int cursorPosition = textBox.SelectionStart;

                    // Pokud je kurzor na začátku textu, smaž první znak
                    if (cursorPosition == 0)
                    {
                        textBox.Text = textBox.Text.Substring(1);
                        textBox.SelectionStart = 0;
                    }
                    // Pokud je kurzor za posledním znakem textu, smaž poslední znak
                    else if (cursorPosition == textBox.Text.Length)
                    {
                        textBox.Text = textBox.Text.Substring(0, textBox.Text.Length - 1);
                        textBox.SelectionStart = cursorPosition - 1;
                    }
                    // Pokud je kurzor uvnitř textu, smaž znak před kurzorem
                    else
                    {
                        // Rozdělení textu na část před kurzorem a část za kurzorem
                        string textBeforeCursor = textBox.Text.Substring(0, cursorPosition - 1);
                        string textAfterCursor = textBox.Text.Substring(cursorPosition);

                        // Aktualizace textu s odstraněním znaku na pozici kurzoru
                        textBox.Text = textBeforeCursor + textAfterCursor;

                        // Nastavení pozice kurzoru po smazání znaku
                        textBox.SelectionStart = cursorPosition - 1;
                    }

                    textBox.ScrollToCaret();
                }
            }


            public static void MoveCursorRight(System.Windows.Forms.TextBox textBox)
            {
                if (textBox == null) return;

                if (textBox.InvokeRequired)
                {
                    textBox.Invoke(new Action(() => MoveCursorRight(textBox)));
                }
                else
                {
                    if (textBox.SelectionStart < textBox.Text.Length)
                    {
                        textBox.SelectionStart++;
                        textBox.ScrollToCaret();
                    }
                }
            }

            public static void MoveCursorLeft(System.Windows.Forms.TextBox textBox)
            {
                if (textBox == null) return;

                if (textBox.InvokeRequired)
                {
                    textBox.Invoke(new Action(() => MoveCursorLeft(textBox)));
                }
                else
                {
                    if (textBox.SelectionStart > 0)
                    {
                        textBox.SelectionStart--;
                        textBox.ScrollToCaret();
                    }
                }
            }

        }


        public void ComReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort serialPort = (SerialPort)sender;
            string data = serialPort.ReadLine();//Rea.ReadExisting(); // Přečtení všech dostupných dat
            serialPort.DiscardInBuffer();
            findText.Invoke((MethodInvoker)delegate {
                findText.Focus();
                findText.Text = "";
                findText.Text = data;
            });
            TextBoxHelper.MoveCursorToEnd(findText);
            find.Invoke((MethodInvoker)delegate {
                find.PerformClick();
            });
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
                MessageBox.Show(ex.Message, @"Problém s uzavřením COM portu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void find_Click(object sender, EventArgs e)
        {
            string textBox = findText.Text;
            textBox = textBox.Replace("\r", "");
            if (textBox.Length == 0)
            {
                Console.Beep(333, 150);
                name.Text = "";
                count.Text = "00000";
                pos.Text = "00000";
                findText.Focus();
            }
            else
            {
                Part tryFind = FindPart(textBox);
                if (tryFind.KZM == "0")
                {
                    List<Part> listParts = FindPartsByAllName(textBox);
                    if (listParts.Count == 0)
                    {
                        Console.Beep(333, 150);
                        name.Text = "";
                        count.Text = "00000";
                        pos.Text = "00000";
                        findText.Focus();
                    }
                    else
                    {
                        listBox1.Items.Clear();
                        for (int i = 0; i < listParts.Count; i++)
                        {
                            listBox1.Items.Add( 
                                "PN: " + listParts[i].PartNumber + " | " +
                                "Název: " + listParts[i].Nazev + " | " +
                                "Místo: " + listParts[i].Umisteni  + " | " +
                                "Počet: " + listParts[i].Pocet);
                        }
                        listBox1.Show();
                        buttoncloselist.Show();
                        listBox1.Focus();
                        listBox1.SelectedIndex = 0;
                    }
                }
                else 
                {
                    // 32.1575.400-06
                    name.Text = tryFind.Nazev;
                    count.Text = tryFind.Pocet;
                    pos.Text = tryFind.Umisteni;
                    findText.Focus();
                    Console.Beep(1233, 80);
                    Console.Beep(1033, 80);
                }
            }
        }

        private void buttoncloselist_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.Hide();
            buttoncloselist.Hide();
            findText.Focus();
        }

        private void EditKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true; // Potlačí zvukový signál nebo jiný efekt
                find_Click(sender, e);
            }
        }


        private string ExtractNazevFromText(string text)
        {
            string prefix = "Název: ";
            int startIndex = text.IndexOf(prefix);
            if (startIndex != -1)
            {
                startIndex += prefix.Length;
                int endIndex = text.IndexOf('|', startIndex);
                if (endIndex == -1)
                {
                    endIndex = text.Length;
                }
                return text.Substring(startIndex, endIndex - startIndex).Trim();
            }
            return string.Empty;
        }

        private string ExtractMistoFromText(string text)
        {
            string prefix = "Místo: ";
            int startIndex = text.IndexOf(prefix);
            if (startIndex != -1)
            {
                startIndex += prefix.Length;
                int endIndex = text.IndexOf('|', startIndex);
                if (endIndex == -1)
                {
                    endIndex = text.Length;
                }
                return text.Substring(startIndex, endIndex - startIndex).Trim();
            }
            return string.Empty;
        }


        private string ExtractPocetFromText(string text)
        {
            string prefix = "Počet: ";
            int startIndex = text.IndexOf(prefix);
            if (startIndex != -1)
            {
                startIndex += prefix.Length;
                int endIndex = text.IndexOf('|', startIndex);
                if (endIndex == -1)
                {
                    endIndex = text.Length;
                }
                return text.Substring(startIndex, endIndex - startIndex).Trim();
            }
            return string.Empty;
        }


        private void PartsListSelectedChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                name.Text = ExtractNazevFromText(listBox1.SelectedItem.ToString());
                count.Text = ExtractPocetFromText(listBox1.SelectedItem.ToString());
                pos.Text = ExtractMistoFromText(listBox1.SelectedItem.ToString());
            }
        }

        private void PartsListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true; // Potlačí zvukový signál nebo jiný efekt
                listBox1.Items.Clear();
                listBox1.Hide();
                buttoncloselist.Hide();
                findText.Focus();
            }
        }

        // smaze cely text v editboxu
        private void delete_Click(object sender, EventArgs e)
        {
            findText.Invoke((MethodInvoker)delegate {
                findText.Text = "";
                findText.Focus();
            });
        }

        // smaze pismenko - stejne jako backspace
        private void backspace_Click(object sender, EventArgs e)
        {
            TextBoxHelper.RemoveCharacterAtCursor(findText);
            findText.Focus();
        }

        // mezernik
        private void bspace_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, " ");
            findText.Focus();
        }

        // sipka vlevo
        private void bleft_Click(object sender, EventArgs e)
        {
            TextBoxHelper.MoveCursorLeft(findText);
            findText.Focus();
        }

        // sipka vpravo
        private void bright_Click(object sender, EventArgs e)
        {
            TextBoxHelper.MoveCursorRight(findText);
            findText.Focus();
        }

        // tecka
        private void bdot_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, ".");
            findText.Focus();
        }

        // pomlcka
        private void bdash_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "-");
            findText.Focus();
        }

        // text 32.
        private void b32_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "32.");
            findText.Focus();
        }

        // text 320.
        private void b320_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "320.");
            findText.Focus();
        }

        private void b1_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "1");
            findText.Focus();
        }

        private void b2_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "2");
            findText.Focus();
        }

        private void b3_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "3");
            findText.Focus();
        }

        private void b4_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "4");
            findText.Focus();
        }

        private void b5_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "5");
            findText.Focus();
        }

        private void b6_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "6");
            findText.Focus();
        }

        private void b7_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "7");
            findText.Focus();
        }

        private void b8_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "8");
            findText.Focus();
        }

        private void b9_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "9");
            findText.Focus();
        }

        private void b0_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "0");
            findText.Focus();
        }

        private void bq_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "Q");
            findText.Focus();
        }

        private void bw_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "W");
            findText.Focus();
        }

        private void be_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "E");
            findText.Focus();
        }

        private void br_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "R");
            findText.Focus();
        }

        private void bt_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "T");
            findText.Focus();
        }

        private void bz_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "Z");
            findText.Focus();
        }

        private void bu_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "U");
            findText.Focus();
        }

        private void bi_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "I");
            findText.Focus();
        }

        private void bo_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "O");
            findText.Focus();
        }

        private void bp_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "P");
            findText.Focus();
        }

        private void ba_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "A");
            findText.Focus();
        }

        private void bs_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "S");
            findText.Focus();
        }

        private void bd_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "D");
            findText.Focus();
        }

        private void bf_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "F");
            findText.Focus();
        }

        private void bg_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "G");
            findText.Focus();
        }

        private void bh_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "H");
            findText.Focus();
        }

        private void bj_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "J");
            findText.Focus();
        }

        private void bk_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "K");
            findText.Focus();
        }

        private void bl_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "L");
            findText.Focus();
        }

        private void by_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "Y");
            findText.Focus();
        }

        private void bx_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "X");
            findText.Focus();
        }

        private void bc_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "C");
            findText.Focus();
        }

        private void bv_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "V");
            findText.Focus();
        }

        private void bb_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "B");
            findText.Focus();
        }

        private void bn_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "N");
            findText.Focus();
        }

        private void bm_Click(object sender, EventArgs e)
        {
            TextBoxHelper.AppendTextAtCursorAndMoveCursorToEnd(findText, "M");
            findText.Focus();
        }
    }
}
