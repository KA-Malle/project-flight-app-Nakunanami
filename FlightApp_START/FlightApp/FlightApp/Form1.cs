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

            int rij = 9;
            string departure_id, departure, arrivel_id, arrivel_city, max_capacity, actual_capacity, date_of_flight,average_cost, type_of_flight;

            // Openen van de werkmap
            xlWerkmap = xlToep.Workbooks.Open(Application.StartupPath + @"\FlightApp.xlsx");
            //Open werkblad
            xlWerkblad = xlWerkmap.ActiveSheet;

            // Keuzelijst leegmaken
            listBoxFlights.Items.Clear();
        }

        private void afsluitenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();

        }
    }
}
