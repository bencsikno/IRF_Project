using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace beadando.Entities
{
    public enum Position
    {
        Irányító= 1,
        Bedobó = 2,
        Kiscsatár = 3,
        Erőcsatár = 4,
        Center = 5
    }

    public class Players
    {
        //public static int Count { get; internal set; }
        public string Nev { get; set; }
        public Position Position { get; set; }
        public double Perc { get; set; }
        public double PontAtlag { get; set; }
        public double DobottPerPerc { get; set; }

    }
}
