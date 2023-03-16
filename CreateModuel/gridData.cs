using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;

namespace TestTekla
{
   public  class gridData
    {
        public int ID { get; set; }

        public string gridName { get; set; }

        public Point startPoint { get; set; }

        public Point endPoint { get; set; }

        //public string startPoint { get; set; }

        //public string endPoint { get; set; }

    }
}
