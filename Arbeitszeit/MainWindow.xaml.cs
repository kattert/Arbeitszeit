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
using ClosedXML.Excel;
using System.IO;

namespace Arbeitszeit
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static bool Läuft;

        public static XLWorkbook workbook;

        public static char[] alphabet = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'V', 'W', 'X', 'Y', 'Z',};

        public static string DatumZeit(int option)
        {
            //Datum oder Zeit zurückgeben. option: 1 = Year; 2 = Month; 3 = Day; 4 = Hour; 5 = Minute; 6 = Secound, 7 = Time, 8 = Date, 9 = Day of Year

            DateTime localDateTime = DateTime.Now;
            string temp;

            switch (option)
            {
                case 1:
                    temp = localDateTime.Year.ToString();
                    return temp;
                case 2:
                    temp = localDateTime.Month.ToString();
                    return temp;
                case 3:
                    temp = localDateTime.Day.ToString();
                    return temp;
                case 4:
                    temp = localDateTime.Hour.ToString();
                    return temp;
                case 5:
                    temp = localDateTime.Minute.ToString();
                    return temp;
                case 6:
                    temp = localDateTime.Second.ToString();
                    return temp;
                case 7:
                    temp = localDateTime.TimeOfDay.ToString();
                    return temp;
                case 8:
                    temp = localDateTime.Date.ToString();
                    return temp;
                case 9:
                    temp = localDateTime.DayOfYear.ToString();
                    return temp;
                default:
                    return null;
            }


        }

        public static bool DateiAnlegen()
        {
            string year = DatumZeit(1);

            string file = "Year_" + year + ".xlsx";

            bool exists = File.Exists(file);

            if (exists)
            {
                workbook  = new XLWorkbook(file);
                return true;
            }
            else
            {
                workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("year");
                worksheet.Cell("A1").Value = "day of year";
                worksheet.Cell("B1").Value = "date";
                worksheet.Cell("C1").Value = "working hours";
                worksheet.Cell("D1").Value = "start";
                worksheet.Cell("E1").Value = "stopp";
                worksheet.Cell("F1").Value = "auftrag";
                worksheet.Cell("G1").Value = "start";
                worksheet.Cell("H1").Value = "stopp";
                worksheet.Cell("I1").Value = "auftrag";
                worksheet.Cell("J1").Value = "start";
                worksheet.Cell("K1").Value = "stopp";
                worksheet.Cell("L1").Value = "auftrag";
                worksheet.Cell("M1").Value = "start";
                worksheet.Cell("N1").Value = "stopp";
                worksheet.Cell("O1").Value = "auftrag";
                worksheet.Cell("P1").Value = "start";
                worksheet.Cell("Q1").Value = "stopp";
                worksheet.Cell("R1").Value = "auftrag";
                worksheet.Cell("S1").Value = "start";
                worksheet.Cell("T1").Value = "stopp";
                worksheet.Cell("U1").Value = "auftrag";
                worksheet.Cell("V1").Value = "start";
                worksheet.Cell("W1").Value = "stopp";
                worksheet.Cell("X1").Value = "auftrag";

                for (int i = 1; i < 367; i++)
                {
                    string zelle = "a" + (i + 1).ToString();
                    worksheet.Cell(zelle).Value = i.ToString();
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(file);
                return true;
            }


        }

        private void StatusErmitteln()
        {
            // Aktuelle Zeit ermitteln
            DateTime localDateTime = DateTime.Now;

            // Zeile in Tabelle ermitteln
            int DayOfYear = localDateTime.DayOfYear;
            int zeilenNummer = DayOfYear + 1;
            // Öffne Worksheet
            var worksheet = workbook.Worksheet("year");

            int startZelle = 3;
            int i = 0;
            bool zeileEmpty = false;
            string zeile;

            while (!zeileEmpty)
            {
                zeile = alphabet[startZelle + i] + zeilenNummer.ToString();
                zeileEmpty = worksheet.Cell(zeile).IsEmpty();
                if (zeileEmpty)
                {
                    Läuft = false;
                }
                else
                {
                    zeile = alphabet[startZelle + i +1] + zeilenNummer.ToString();
                    zeileEmpty = worksheet.Cell(zeile).IsEmpty();
                    if (zeileEmpty)
                    {
                        Läuft = true;
                    }
                }
                if (!zeileEmpty)
                {
                    i += 3;
                }
            }
            if (Läuft)
            {
                buttonStempeln.Background = Brushes.Green;
            }
            else
            {
                buttonStempeln.Background = Brushes.Gray;
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            DateiAnlegen();
            StatusErmitteln();
        }

        private void buttonStempeln_Click(object sender, RoutedEventArgs e)
        {
            // Aktuelle Zeit ermitteln
            DateTime localDateTime = DateTime.Now;

            // Zeile in Tabelle ermitteln
            int DayOfYear = localDateTime.DayOfYear;
            int zeilenNummer = DayOfYear + 1;

            // Öffne Worksheet
            var worksheet = workbook.Worksheet("year");

            // Datum eintragen
            worksheet.Cell("B" + zeilenNummer.ToString()).Value = localDateTime.Date.ToString("d");

            // Eingestempelt oder ausgestempelt
            if (Läuft)
            {
                Läuft = false;
                buttonStempeln.Background = Brushes.Gray;
            }
            else
            {
                Läuft = true;
                buttonStempeln.Background = Brushes.Green;
            }

            string zeile;
            bool zeileEmpty;
            bool eingetragen = false;
            int i = 0;
            int startZelle = 3;

            // Start und Stopp Zeiten eintragen
            while (!eingetragen)
            {
                if (Läuft)
                {
                    zeile = alphabet[startZelle + i] + zeilenNummer.ToString();
                    zeileEmpty = worksheet.Cell(zeile).IsEmpty();
                    if (zeileEmpty)
                    {
                        worksheet.Cell(zeile).Value = localDateTime.TimeOfDay;
                        zeile = alphabet[startZelle + i + 2] + zeilenNummer.ToString();
                        worksheet.Cell(zeile).Value = textBoxAuftrag.Text;
                        eingetragen = true;
                    }
                }
                else
                {
                    zeile = alphabet[startZelle + i + 1] + zeilenNummer.ToString();
                    string zeile2 = alphabet[startZelle + i + 3] + zeilenNummer.ToString();
                    zeileEmpty = worksheet.Cell(zeile).IsEmpty() & worksheet.Cell(zeile2).IsEmpty();
                    if (zeileEmpty)
                    {
                        worksheet.Cell(zeile).Value = localDateTime.TimeOfDay;
                        eingetragen = true;
                    }
                }

                if (!eingetragen)
                {
                    i += 3;
                }


            }

            // Arbeitszeit ermitteln
            if (!Läuft)
            {
                i = 0;
                DateTime summeZeit = new DateTime();
                zeile = alphabet[startZelle] + zeilenNummer.ToString();
                bool zeileFilled = !worksheet.Cell(zeile).IsEmpty();

                while (zeileFilled)
                {
                    zeile = alphabet[startZelle + i] + zeilenNummer.ToString();
                    zeileFilled = !worksheet.Cell(zeile).IsEmpty();
                    if (zeileFilled)
                    {
                        DateTime wert1 = worksheet.Cell(zeile).GetValue<DateTime>();
                        zeile = alphabet[startZelle + 1 + i] + zeilenNummer.ToString();
                        zeileFilled = !worksheet.Cell(zeile).IsEmpty();
                        if (zeileFilled)
                        {
                            DateTime wert2 = worksheet.Cell(zeile).GetValue<DateTime>();
                            Console.WriteLine(wert2 - wert1);
                            summeZeit += (wert2 - wert1);
                            Console.WriteLine("Summe: " + summeZeit);

                            i += 3;
                        }
                    }
                }
                worksheet.Cell("C" + zeilenNummer.ToString()).Value = summeZeit;
                textBoxArbeitszeit.Text = summeZeit.ToString("T");
            }

            // Workbook Speichern
            workbook.Save();
        }

        private void buttonBuchungen_Click(object sender, RoutedEventArgs e)
        {
            DateTime localDateTime = DateTime.Now;

            // Zeile in Tabelle ermitteln
            int DayOfYear = localDateTime.DayOfYear;
            int zeilenNummer = DayOfYear + 1;

            // Öffne Worksheet
            var worksheet = workbook.Worksheet("year");

            string text;

            if (!worksheet.Cell("B" + zeilenNummer.ToString()).IsEmpty())
            {
                string datum = worksheet.Cell("B" + zeilenNummer.ToString()).Value.ToString();
                text = "Datum: " + datum.Substring(0, 10) + "\n";
                text += "\n";
                text += "      Start        Stopp     Auftrag" + "\n";
               

                int i = 0;
                bool zeileEmpty = false;
                string zeile;
                int startZelle = 3;

                while (!zeileEmpty)
                {
                    zeile = alphabet[startZelle + i] + zeilenNummer.ToString();
                    zeileEmpty = worksheet.Cell(zeile).IsEmpty();
                    if (!zeileEmpty)
                    {
                        text += "   " + worksheet.Cell(zeile).Value.ToString().Substring(0,8);

                        zeile = alphabet[startZelle + i + 1] + zeilenNummer.ToString();
                        zeileEmpty = worksheet.Cell(zeile).IsEmpty();
                        if (!zeileEmpty)
                        {
                            text += "   " + worksheet.Cell(zeile).Value.ToString().Substring(0, 8);

                            zeile = alphabet[startZelle + i + 2] + zeilenNummer.ToString();
                            zeileEmpty = worksheet.Cell(zeile).IsEmpty();
                            if (!zeileEmpty)
                            {
                                text += "   " + worksheet.Cell(zeile).Value.ToString();
                                
                                i += 3;
                            }
                        }
                    }
                    text += "\n";

                }

                if (!worksheet.Cell("C" + zeilenNummer.ToString()).IsEmpty())
                {
                    string arbeitszeit = worksheet.Cell("C" + zeilenNummer.ToString()).Value.ToString();
                    text += "\n";
                    text += "Arbeitszeit: " + arbeitszeit.Substring(11);
                }
                else
                {
                    text += "Arbeitszeit: " + "0,0";
                    MessageBox.Show("Heute noch nichts geleistet?");
                }
                richTextBoxBuchungen.Document.Blocks.Clear();
                richTextBoxBuchungen.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
            else
            {
                MessageBox.Show("Heute noch nichts gebucht?");
            }
        }
    }
}
