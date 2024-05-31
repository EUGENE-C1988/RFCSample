using SapNwRfc;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetCoreSAPSample.Model
{
    internal class ZMMI0025
    {
        internal class Param
        {
            [SapName("IS_PO")]
            public string ZTYPE { get; set; }
            public string EEIND_F { get; set; }
            public string EEIND_T { get; set; }
        }

        internal class ET_PO_result
        {
            [SapName("ET_PO")]
            public ET_PO_Item[] ItemList { get; set; }
        }
        public class ET_PO_Item
        {
            public string WERKS { get; set; }
            public string EBELN { get; set; }
            public string MATNR { get; set; }
            public string MAKTX { get; set; }
        }
    }
}
