using SapNwRfc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DotNetCoreSAPSample.Model.ZPPI0008;

namespace DotNetCoreSAPSample.Model
{
    internal class ZPPI0008
    {
        internal class Param
        {
            [SapName("I_ALL")]
            public string I_ALL { get; set; }
        }

        internal class ZMMI0008r
        {
            [SapName("O_ZPPI0008")]
            public ZPPI0008_Item[] ItemList { get; set; }
        }
        public class ZPPI0008_Item
        {
            public string MATNR { get; set; }
            public string VORNR { get; set; }
            public string ZSEQ { get; set; }
            public string ZPROCESS { get; set; }
            public string ZPROCESSNA1 { get; set; }
            public string ZPLNAL { get; set; }
        }
    }
}
