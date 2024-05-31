using DotNetCoreSAPSample;
using DotNetCoreSAPSample.Model;
using System.Text;

var connectionString = "AppServerHost=192.168.101.26; User=MESRFC; Password=Gpi6128$; SystemNumber=00; Client=667;  SystemId=S4D; Language=ZF; PoolSize=20; Trace=8;";
var sapService = new SAPService(connectionString);
sapService.CallZMMI0025Test();
