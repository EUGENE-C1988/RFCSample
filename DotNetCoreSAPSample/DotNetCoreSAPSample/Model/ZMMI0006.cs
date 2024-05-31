using SapNwRfc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DotNetCoreSAPSample.Model.ZMMI0006;

namespace DotNetCoreSAPSample.Model
{
    internal class ZMMI0006
    {
        internal class Param
        {
            [SapName("I_MATNR")]
            public string I_MATNR { get; set; }
            public string I_WERKS { get; set; }
        }

        internal class ZMMI0006r
        {
            [SapName("ZMMI0006")]
            public ZMMI0006_Item[] ItemList { get; set; }
        }
        public class ZMMI0006_Item
        {
            public string WERKS { get; set; }
            public string MATNR { get; set; }
            public string MAKTX { get; set; }
            public string MATKL { get; set; }

        }
    }
}