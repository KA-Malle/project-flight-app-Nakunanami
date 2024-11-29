using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace FlightApp
{
    public partial class FormFlightApp : Form
    {
        /*
         * Naam: Roxy Sluyts
         * Klas: 6ADB
         */
        public FormFlightApp()
        {
            InitializeComponent();
        }

        // Declaratie
        List<string[]> _flightsList = new List<string[]>();
        List<string[]> _filteredFlightsList = new List<string[]>();


        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Vergeet niet via "add refrence" het Microsoft Excel toe te voegen
            // Declaratie
            Microsoft.Office.Interop.Excel.Application xlToep = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWerkmap;
            Microsoft.Office.Interop.Excel.Worksheet xlWerkblad;


            // Instellen van het dialoogvenster Openen
            OpenFileDialog dlgOpen = new OpenFileDialog();

            //Eigenschappen instellen
            dlgOpen.Title = "Openen";
            dlgOpen.FileName = "";
            dlgOpen.DefaultExt = ".xlsx";
            dlgOpen.InitialDirectory = Application.StartupPath;
            dlgOpen.Filter = "Excel (.xlsx) |*.xlsx|Alle bestanden (*.*)|*.*";

            //Dialoogvenster tonen en de keuze opvangen
            DialogResult resultaat = dlgOpen.ShowDialog();

            //Kijken of de gebruiker "openen" (ok) geklikt heeft
            if (resultaat == DialogResult.OK)
            {
                // Open de werkmap
                xlWerkmap = xlToep.Workbooks.Open(dlgOpen.FileName);
                xlWerkblad = xlWerkmap.Sheets[1]; 

                int rij = 2;

                _flightsList.Clear(); 

                // Doorloop het werkblad
                while (xlWerkblad.Cells[rij, 1].Value != null)
                {
                    // Haal de waarden uit de cellen
                    string date_of_flight = xlWerkblad.Cells[rij, 7].Value.ToString(); // 'yyyy/mm/dd'
                    string departure = xlWerkblad.Cells[rij, 2].Value.ToString(); 
                    string arrivel_city = xlWerkblad.Cells[rij, 4].Value.ToString(); 
                    string actual_capacity = xlWerkblad.Cells[rij, 6].Value.ToString();
                    string max_capacity = xlWerkblad.Cells[rij, 7].Value.ToString();
                    string type_of_flight = xlWerkblad.Cells[rij, 9].Value.ToString();

                    
 
                    // Voeg de gegevens toe aan de lijst
                    _flightsList.Add(new string[] {date_of_flight, departure, arrivel_city, actual_capacity, max_capacity, type_of_flight });

                    rij++; 
                }

                // Clone the list
                _filteredFlightsList = _flightsList.Select(flight => (string[])flight.Clone()).ToList();

                // Show the list in the ListBox
                ShowInList(_filteredFlightsList);

                // Excel afsluiten
                xlToep.Quit();

            }
        }

        private void ShowInList(List<string[]> flightsList)
        {
            // ListBox leegmaken
            //listBoxFlights.Items.Clear(); 

            foreach (var flight in _filteredFlightsList)
            {
                // Declaratie 
                string temp = "";

                //datum, departure, arrival, atual and max capacity, type

                // Tekst in label opbouwen met padding
                temp = (flight[0]).PadRight(5);
                temp += flight[1].PadRight(15) + " =>" + flight[2].PadRight(20);
                temp += (flight[3] + "/" + flight[4]).PadRight(10);
                temp += (flight[5]).PadRight(10);

                listBoxFlights.Items.Add(temp);
            }
        }

        private void afsluitenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();

        }
    }
}
