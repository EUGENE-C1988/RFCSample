using SapNwRfc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DotNetCoreSAPSample.Model.ZPPI0003;

namespace DotNetCoreSAPSample.Model
{
    internal class ZPPI0003
    {
        //輸入參數
        internal class Param
        {
            [SapName("I_AUFNR")]
            public string I_AUFNR { get; set; }
        }

        //回傳參數物件型態
        internal class ZPPI0002
        {
            [SapName("ZPPI0002")]
            public ZPPI0002_Item[] ItemList { get; set; }
        }

        internal class ZPPI0002_BOM
        {
            [SapName("ZPPI0002_BOM")]
            public ZPPI0002_BOM_Item[] ItemList { get; set; }
        }

        internal class ZPPI0002_ROUT
        {
            [SapName("ZPPI0002_ROUT")]
            public ZPPI0002_ROUT_Item[] ItemList { get; set; }
        }

        public class ZPPI0002_Item
        {
            public string AUFNR { get; set; }
            public string TECO_FLAG { get; set; }
            public string DLV_FLAG { get; set; }
            public string CLSD_FLAG { get; set; }
            public string DLT_FLAG { get; set; }

        }

        public class ZPPI0002_BOM_Item
        {
            public string FMATNR { get; set; }
            public string FMAKTX { get; set; }
            public string MATKL { get; set; }
            public string WGBEZ { get; set; }
            
        }

        public class ZPPI0002_ROUT_Item
        {
            public string VORNR { get; set; }
        }
    }

}
