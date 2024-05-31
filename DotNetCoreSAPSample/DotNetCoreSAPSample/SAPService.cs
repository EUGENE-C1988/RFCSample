using DotNetCoreSAPSample.Model;
using SapNwRfc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetCoreSAPSample
{
    internal class SAPService
    {
        private SapConnection sapConnection;

        public SAPService(string connectionString)
        {
            try
            {
                sapConnection = new SapConnection(connectionString);
                sapConnection.Connect();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CallZPPI0003Test()
        {
            using var sapFunctionClient = sapConnection.CreateFunction("ZPPI0003");
            var result = sapFunctionClient.Invoke<ZPPI0003.ZPPI0002>(new ZPPI0003.Param { I_AUFNR = "100176522" });
            
        }
        public void CallZMMI0025Test()
        {
            using var sapFunctionClient = sapConnection.CreateFunction("ZMMI0025");
            //var Reslut = sapFunctionClient.Invoke<ZPPI0008.ZMMI0008r>(new ZPPI0008.Param { I_ALL = "X"});
            var result = sapFunctionClient.Invoke<ZMMI0025.ET_PO_result>(new ZMMI0025.Param { ZTYPE = "1", EEIND_F="20230502", EEIND_T="20230513" });

        }
    }
}
