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
using System.Drawing.Text;
using Spire.Xls;
using Spire.Xls.Collections;

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

            RemoveColoumn(xlWorkSheet,"A1");

            AddRow(xlWorkSheet,1);

            Formazas(xlWorkSheet);
            JatekosMinoseg(xlWorkSheet);


        }
        private void RemoveColoumn(Excel.Worksheet xlWorkSheet, string coloumnname) 
        {
            Excel.Range range = (Excel.Range)xlWorkSheet.get_Range(coloumnname, Missing.Value);
            range.EntireColumn.Delete(Missing.Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);            

        }
        private void AddRow(Excel.Worksheet xlWorkSheet, int rowindex)
        {
            Excel.Range range = (Excel.Range)xlWorkSheet.Rows[rowindex];
            range.Insert();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

            Fejlecek(xlWorkSheet);

        }
        private void Fejlecek(Excel.Worksheet xlWorkSheet) 
        {
            xlWorkSheet.Cells[1, 1] = "Név";
            xlWorkSheet.Cells[1, 2] = "Pozíció";
            xlWorkSheet.Cells[1, 3] = "Játszott perc";
            xlWorkSheet.Cells[1, 4] = "Pontátlag";
            xlWorkSheet.Cells[1, 5] = "Percenkénti pontok";

        }

        private void Formazas(Excel.Worksheet xlWorkSheet) 
        {
            Excel.Range headerRange = xlWorkSheet.get_Range("A1","E1");
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.Fuchsia;
            Excel.Range lastsor = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range totalRange = xlWorkSheet.get_Range("A2", "E"+lastsor.Row);

            totalRange.BorderAround2(Excel.XlLineStyle.xlDash, Excel.XlBorderWeight.xlMedium);
            totalRange.Cells.Borders.LineStyle = Excel.XlLineStyle.xlDashDot;
            totalRange.Interior.Color = Color.Yellow;
                

            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);




        }

        private void JatekosMinoseg(Excel.Worksheet xlWorkSheet) 
        {

            Excel.Range lastsor = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range totalRange = xlWorkSheet.get_Range("A2", "E" + lastsor.Row);

            totalRange.Sort(totalRange.Columns[5], Excel.XlSortOrder.xlDescending);


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
