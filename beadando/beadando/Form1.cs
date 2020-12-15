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
