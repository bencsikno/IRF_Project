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
using System.Diagnostics;

namespace beadando
{
    public partial class Form1 : Form
    {
        List<Players> Jatekosok = new List<Players>();        
        public Form1()
        {
            InitializeComponent();
            Jatekosok = GetPlayers(AppDomain.CurrentDomain.BaseDirectory + @"\jatekosadatok\nbajatekosok.csv");
            dataGridView1.DataSource = Jatekosok.ToList();
           


        }
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
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

        private void button1_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void Szures_Click(object sender, EventArgs e)
        {
            List<Players> Filteredplayers = new List<Players>();

            int min = (int)numericUpDown1.Value;
            string pozi = ""; 
            if (comboBox1.SelectedItem != null)
            {
                pozi = comboBox1.SelectedItem.ToString();
            }
                

            //Debug.WriteLine(pozi);
            //Debug.WriteLine(Position.Erőcsatár.ToString());
            for (int i = 0; i < Jatekosok.Count; i++)
            {
                if (Jatekosok[i].Perc>=min)
                {
                    if ((pozi!="") && (Jatekosok[i].Position.ToString() == pozi) )
                    {
                        Filteredplayers.Add(Jatekosok[i]);
                    }
                    else if (pozi=="")
                    {
                        Filteredplayers.Add(Jatekosok[i]);
                    }                 
                     
                    
                   

                }
            }                      

            dataGridView1.DataSource = Filteredplayers.ToList();
        }
    }
}
