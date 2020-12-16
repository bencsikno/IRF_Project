using beadando.Entities;
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
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace beadando
{
    public partial class Form1 : Form
    {
        List<Players> Jatekosok = new List<Players>();


        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;
        public Form1()
        {
            InitializeComponent();
            Jatekosok = GetPlayers(AppDomain.CurrentDomain.BaseDirectory + @"\jatekosadatok\nbajatekosok.csv");
            dataGridView1.DataSource = Jatekosok.ToList();
            CreateExcel();

            CreateTable();

        }

        private static void CreateTable()
        {
            string[] headers = new string[]
            {
                "Név",
                "Pozíció",
                "Perc",
                "Pontátlag",
                "Percenként dobott pontátlag"
            };
            for (int i = 0; i < 290; i++)
            {
                xlSheet.Cells[1, 1] = headers[0];
                object[,] values = new object[, headers.Length];
            }
            int counter = 0;
            foreach (Players f in Jatekosok)
            {
                values[counter, 0] = f.Code;
                // ...
                values[counter, 8] = "";
                counter++;
            }
        }

        private void CreateExcel()
        {
            try
            {

                xlApp = new Excel.Application();

                xlWB = xlApp.Workbooks.Add(Missing.Value);

                xlSheet = xlWB.ActiveSheet;


                CreateTable();


                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        public List<Players> GetPlayers(string csvpath)
        {
            List<Players> players = new List<Players>();
            using (StreamReader sr = new StreamReader(csvpath, Encoding.Default))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine().Split(';');
                    players.Add(new Players()
                    {
                        Nev = (line[0]),
                        Position = (Position)Enum.Parse(typeof(Position), line[1]),
                        Perc = double.Parse(line[2]),
                        PontAtlag = double.Parse(line[3]),
                        DobottPerPerc = double.Parse(line[4])
                    } );
                }

                return players;
            }
        }

        








    }
}
