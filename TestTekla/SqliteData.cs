using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestTekla
{
   public  class SqliteData
    {
        public string ProfileName { get;set; }
        public double AverageRate { get; set; }
        public string MaterialName { get; set; }
        public double Number { get; set; }

        public string Other { get; set; }

        public double  MateriAllNumber{get;set;}

        public double Rate { get; set; }

        public double MateriAllLength { get; set; }
        public double WasteLength { get; set; }

        public string LengthList { get; set; }
        public string IdList { get; set; }

        public string WeightList { get; set; }

        public string MateriList { get; set; }

        public bool IsChecked { get; set; }

    }
}
