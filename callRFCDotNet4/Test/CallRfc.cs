using System.Linq;
using System.Web;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using SAP.Middleware.Connector;
using System.IO;
using System.Data.OleDb;

namespace CallRfc
{
    public class RFC_Tool : IDisposable
    {
        protected RfcConfigParameters parms;

        protected RfcDestination rfcDest;

        protected RfcRepository rfcRep;

        protected IRfcFunction function;

        protected Action<string, string, int> ActionString;

        protected Action<string, int, int> ActionInt;

        protected Action<string, long, int> Actionlong;

        protected Action<string, char, int> ActionChar;

        protected Action<string, decimal, int> ActionDecimal;

        protected Action<string, double, int> ActionDouble;

        protected Action<string, float, int> ActionSingle;

        protected Action<string, DateTime, int> ActionDateTime;

        public WriteToFile WR = new WriteToFile();

        protected List<InputParameter> ParList = new List<InputParameter>();

        protected List<OutParameter> OutParList = new List<OutParameter>();

        private bool disposed = false;

        private SafeHandle handle = new SafeFileHandle(IntPtr.Zero, ownsHandle: true);

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    handle.Dispose();
                    parms = null;
                    rfcDest = null;
                    rfcRep = null;
                    function = null;
                    ClearInputParameterList();
                    ClearOutParameterList();
                    WR.DisconnectToDB();
                    WR = null;
                }
                disposed = true;
            }
        }

        protected RFC_Tool()
        {
        }

        public RFC_Tool(string FunctionName, out bool Sucess)
        {
            Sucess = RFC_Init(FunctionName);
        }

        public RFC_Tool(out bool Sucess)
        {
            Sucess = RFC_Init(null);
        }

        protected bool RFC_Init(string FunctionName)
        {
            try
            {
                WR.WriteDebugLog("***RFC Begin Init");
                parms = new RfcConfigParameters();
                parms.Add("NAME", ConfigurationManager.AppSettings["RFC_Name"]);
                parms.Add("ASHOST", ConfigurationManager.AppSettings["RFC_AppServerHost"]);
                parms.Add("USER", ConfigurationManager.AppSettings["RFC_User"]);
                parms.Add("PASSWD", ConfigurationManager.AppSettings["RFC_Password"]);
                parms.Add("SYSNR", ConfigurationManager.AppSettings["RFC_SystemNumber"]);
                parms.Add("CLIENT", ConfigurationManager.AppSettings["RFC_Client"]);
                parms.Add("LANG", ConfigurationManager.AppSettings["RFC_Language"]);
                parms.Add("POOL_SIZE", ConfigurationManager.AppSettings["RFC_PoolSize"]);
                parms.Add("MAX_POOL_WAIT_TIME", ConfigurationManager.AppSettings["RFC_MaxPoolWaitTime"]);
                parms.Add("IDLE_TIMEOUT", ConfigurationManager.AppSettings["RFC_IdleTimeout"]);
                parms.Add("SYSID", ConfigurationManager.AppSettings["RFC_SystemID"]);
                rfcDest = RfcDestinationManager.GetDestination(parms);
                rfcRep = rfcDest.Repository;
                SetRFCFunctionName(FunctionName);
                WR.WriteDebugLog("***RFC connects sucess!");
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                WR.WriteToLogFile("***RFC Init Fail!");
                return false;
            }
        }

        public void SetRFCFunctionName(string FunctionName)
        {
            function = rfcRep.CreateFunction(FunctionName);
        }

        public void ClearInputParameterList()
        {
            for (int i = 0; i < ParList.Count; i++)
            {
                ParList[i].Dispose();
            }
            ParList.Clear();
        }

        public void ClearOutParameterList()
        {
            for (int i = 0; i < OutParList.Count; i++)
            {
                OutParList[i].Dispose();
            }
            OutParList.Clear();
        }

        public bool SetInputParameter<T>(string Name, T value)
        {
            try
            {
                dynamic val = value;
                ParList.Add(new InputParameter(Name, val));
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool SetOutParameter(string Name, OutParameter.OutType type)
        {
            try
            {
                OutParList.Add(new OutParameter(Name, type));
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool GetReturnData(string ParameterName, ref DataTable TemprfcTable)
        {
            WR.WriteDebugLog("***RFC GetReturnData Init Table");
            TemprfcTable.Reset();
            for (int i = 0; i < OutParList.Count; i++)
            {
                if (OutParList[i].ParameterName.Equals(ParameterName))
                {
                    WR.WriteDebugLog("***RFC GetReturnData Table execute");
                    TemprfcTable = OutParList[i].DTable;
                    return true;
                }
            }
            WR.WriteDebugLog("***RFC GetReturnData Table Fail");
            return false;
        }

        public DataTable GetReturnData(string ParameterName)
        {
            WR.WriteDebugLog("***RFC GetReturnData Init Table");
            for (int i = 0; i < OutParList.Count; i++)
            {
                if (OutParList[i].ParameterName.Equals(ParameterName))
                {
                    WR.WriteDebugLog("***RFC GetReturnData Table execute:" + ParameterName);
                    return OutParList[i].DTable;
                }
            }
            WR.WriteDebugLog("***RFC GetReturnData Table Fail");
            return null;
        }

        public bool GetReturnData<T>(string ParameterName, ref T value)
        {
            WR.WriteDebugLog("***RFC GetReturnData Init Single");
            for (int i = 0; i < OutParList.Count; i++)
            {
                if (OutParList[i].ParameterName.Equals(ParameterName))
                {
                    WR.WriteDebugLog("***RFC GetReturnData Single execute");
                    return GetRightType(OutParList[i].Value, ref value);
                }
            }
            WR.WriteDebugLog("***RFC GetReturnData Single Fail");
            return false;
        }

        private bool GetRightType<T>(dynamic Pvalue, ref T value)
        {
            try
            {
                value = Convert.ChangeType(Pvalue, value.GetType());
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public List<T> DatatTableToList<T>(DataTable dt) where T : class, new()
        {
            List<T> list = new List<T>();
            T val = new T();
            PropertyInfo[] properties = val.GetType().GetProperties();
            foreach (DataRow row in dt.Rows)
            {
                val = new T();
                PropertyInfo[] array = properties;
                foreach (PropertyInfo propertyInfo in array)
                {
                    if (dt.Columns.Contains(propertyInfo.Name) && row[propertyInfo.Name] != DBNull.Value)
                    {
                        object value = Convert.ChangeType(row[propertyInfo.Name], propertyInfo.PropertyType);
                        propertyInfo.SetValue(val, value, null);
                    }
                }
                list.Add(val);
            }
            return list;
        }

        public virtual bool AccessViaRFC()
        {
            try
            {
                WR.WriteDebugLog("***RFC Begin AccessViaRFC");
                if (ReadInputPar() && ExecuterfcFunction())
                {
                    ReadOutPar();
                    WR.WriteDebugLog("***RFC Begin AccessViaRFC sucess");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***RFC AccessViaRFC Fail!");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        protected virtual bool ReadInputPar()
        {
            try
            {
                ActionString = delegate (string strName, string Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                ActionInt = delegate (string strName, int Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                Actionlong = delegate (string strName, long Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                ActionChar = delegate (string strName, char Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                ActionDecimal = delegate (string strName, decimal Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                ActionDouble = delegate (string strName, double Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                ActionSingle = delegate (string strName, float Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                ActionDateTime = delegate (string strName, DateTime Value, int index)
                {
                    function.SetValue(strName, Value);
                };
                for (int i = 0; i < ParList.Count; i++)
                {
                    SetValueToRfc(ParList[i].ParameterName, ParList[i].Value, ParList[i].ValueType.Name);
                }
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***RFC ReadInputPar Fail!");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        protected virtual bool ReadOutPar()
        {
            try
            {
                WR.WriteDebugLog("***RFC AccessViaRFC 執行取回 回傳值");
                for (int i = 0; i < OutParList.Count; i++)
                {
                    WR.WriteDebugLog("***RFC AccessViaRFC 執行取回 ParType:" + OutParList[i].ParameterType);
                    if (OutParList[i].ParameterType == OutParameter.OutType.Table)
                    {
                        WR.WriteDebugLog("  ***RFC AccessViaRFC 執行取回 Table:" + OutParList[i].ParameterName);
                        WR.WriteDebugLog("      ITable:" + function.GetTable(OutParList[i].ParameterName).RowCount);
                        DataTable dataTableFromRfcTable = GetDataTableFromRfcTable(function.GetTable(OutParList[i].ParameterName));
                        if (dataTableFromRfcTable != null)
                        {
                            WR.WriteDebugLog("  ***RFC AccessViaRFC 執行取回 ReadOutPar的DataTable:" + OutParList[i].ParameterName + "  Count:" + dataTableFromRfcTable.Rows.Count);
                            OutParList[i].Set_rfcTable(dataTableFromRfcTable);
                            WR.WriteDebugLog("  ***RFC AccessViaRFC 執行取回 ReadOutPar回寫至class的DataTable:" + OutParList[i].ParameterName + "  Count:" + OutParList[i].DTable.Rows.Count);
                        }
                    }
                    else if (OutParList[i].ParameterType == OutParameter.OutType.Structure)
                    {
                        WR.WriteDebugLog("  ***RFC AccessViaRFC 執行取回 Structure:" + OutParList[i].ParameterName);
                        WR.WriteDebugLog("      IStructure:" + function.GetStructure(OutParList[i].ParameterName).Count);
                        DataTable dataStructureFromRfcStructure = GetDataStructureFromRfcStructure(function.GetStructure(OutParList[i].ParameterName));
                        if (dataStructureFromRfcStructure != null)
                        {
                            WR.WriteDebugLog("  ***RFC AccessViaRFC 執行取回 ReadOutPar的DataTable(Structure):" + OutParList[i].ParameterName + "  Count:" + dataStructureFromRfcStructure.Rows.Count);
                            OutParList[i].Set_rfcTable(dataStructureFromRfcStructure);
                            WR.WriteDebugLog("  ***RFC AccessViaRFC 執行取回 ReadOutPar回寫至class的DataTable(Structure):" + OutParList[i].ParameterName + "  Count:" + OutParList[i].DTable.Rows.Count);
                        }
                    }
                    else
                    {
                        OutParList[i].Set_Value(function.GetValue(OutParList[i].ParameterName).ToString());
                        string text = Convert.ToString(OutParList[i].Value);
                        WR.WriteDebugLog("***RFC AccessViaRFC 執行取回 ReadOutPar的Single:" + text);
                    }
                }
                WR.WriteDebugLog("***RFC AccessViaRFC 執行取回 回傳值完成");
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***RFC ReadInputPar(RFC_Tool) Fail");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        protected virtual bool ExecuterfcFunction()
        {
            try
            {
                WR.WriteDebugLog("***RFC AccessViaRFC 執行Invoke");
                function.Invoke(rfcDest);
                WR.WriteDebugLog("***RFC AccessViaRFC 執行Invoke sucess");
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***RFC ExecuterfcFunction Fail");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        protected Type GetDataType(RfcDataType rfcDataType)
        {
            switch (rfcDataType)
            {
                case RfcDataType.DATE:
                case RfcDataType.CHAR:
                case RfcDataType.STRING:
                    return typeof(string);

                case RfcDataType.BCD:
                    return typeof(decimal);

                case RfcDataType.INT2:
                case RfcDataType.INT4:
                    return typeof(int);
                case RfcDataType.FLOAT:
                    return typeof(double);
                default:
                    return typeof(string);
            }
        }

        private DataTable GetDataTableFromRfcTable(IRfcTable rfcTable)
        {
            if (rfcTable == null)
            {
                WR.WriteDebugLog("***RFC GetDataTableFromRfcTable--> rfcTable無資料!");
                return null;
            }
            DataTable dataTable = new DataTable();
            try
            {
                WR.WriteDebugLog("***RFC GetDataTableFromRfcTable 開始");
                WR.WriteDebugLog("***RFC GetDataTableFromRfcTable 第一階段");
                for (int i = 0; i < rfcTable.ElementCount; i++)
                {
                    RfcElementMetadata elementMetadata = rfcTable.GetElementMetadata(i);
                    dataTable.Columns.Add(elementMetadata.Name, GetDataType(elementMetadata.DataType));
                }
                WR.WriteDebugLog("***RFC GetDataTableFromRfcTable 第一階段結束");
                WR.WriteDebugLog("***RFC GetDataTableFromRfcTable 第二階段");
                foreach (IRfcStructure item in rfcTable)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int j = 0; j < rfcTable.ElementCount; j++)
                    {
                        RfcElementMetadata elementMetadata2 = rfcTable.GetElementMetadata(j);
                        switch (elementMetadata2.DataType)
                        {
                            case RfcDataType.DATE:
                                dataRow[elementMetadata2.Name] = item.GetString(elementMetadata2.Name);
                                break;
                            case RfcDataType.BCD:
                                dataRow[elementMetadata2.Name] = item.GetDecimal(elementMetadata2.Name);
                                break;
                            case RfcDataType.CHAR:
                                dataRow[elementMetadata2.Name] = item.GetString(elementMetadata2.Name);
                                break;
                            case RfcDataType.STRING:
                                dataRow[elementMetadata2.Name] = item.GetString(elementMetadata2.Name);
                                break;
                            case RfcDataType.INT2:
                                dataRow[elementMetadata2.Name] = item.GetInt(elementMetadata2.Name);
                                break;
                            case RfcDataType.INT4:
                                dataRow[elementMetadata2.Name] = item.GetInt(elementMetadata2.Name);
                                break;
                            case RfcDataType.FLOAT:
                                dataRow[elementMetadata2.Name] = item.GetDouble(elementMetadata2.Name);
                                break;
                            default:
                                dataRow[elementMetadata2.Name] = item.GetString(elementMetadata2.Name);
                                break;
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
                WR.WriteDebugLog("***RFC GetDataTableFromRfcTable 第二階段結束");
                WR.WriteDebugLog("***RFC GetDataTableFromRfcTable 結束");
                return dataTable;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return null;
            }
        }

        private DataTable GetDataStructureFromRfcStructure(IRfcStructure rfcStructure)
        {
            if (rfcStructure == null)
            {
                WR.WriteDebugLog("***RFC GetDataStructureFromRfcStructure--> IRfcStructure無資料!");
                return null;
            }
            try
            {
                WR.WriteDebugLog("***RFC GetDataStructureFromRfcStructure 開始");
                WR.WriteDebugLog("***RFC GetDataStructureFromRfcStructure 第一階段");
                DataTable dataTable = new DataTable();
                for (int i = 0; i < rfcStructure.ElementCount; i++)
                {
                    RfcElementMetadata elementMetadata = rfcStructure.GetElementMetadata(i);
                    dataTable.Columns.Add(elementMetadata.Name, GetDataType(elementMetadata.DataType));
                }
                WR.WriteDebugLog("***RFC GetDataStructureFromRfcStructure 第一階段結束");
                WR.WriteDebugLog("***RFC GetDataStructureFromRfcStructure 第二階段");
                DataRow dataRow = dataTable.NewRow();
                for (int j = 0; j < rfcStructure.ElementCount; j++)
                {
                    RfcElementMetadata elementMetadata2 = rfcStructure.GetElementMetadata(j);
                    switch (elementMetadata2.DataType)
                    {
                        case RfcDataType.DATE:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetString(elementMetadata2.Name);
                            break;
                        case RfcDataType.BCD:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetDecimal(elementMetadata2.Name);
                            break;
                        case RfcDataType.CHAR:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetString(elementMetadata2.Name);
                            break;
                        case RfcDataType.STRING:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetString(elementMetadata2.Name);
                            break;
                        case RfcDataType.INT2:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetInt(elementMetadata2.Name);
                            break;
                        case RfcDataType.INT4:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetInt(elementMetadata2.Name);
                            break;
                        case RfcDataType.FLOAT:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetDouble(elementMetadata2.Name);
                            break;
                        default:
                            dataRow[elementMetadata2.Name] = rfcStructure.GetString(elementMetadata2.Name);
                            break;
                    }
                }
                dataTable.Rows.Add(dataRow);
                WR.WriteDebugLog("***RFC GetDataStructureFromRfcStructure 第二階段結束");
                WR.WriteDebugLog("***RFC GetDataStructureFromRfcStructure 結束");
                return dataTable;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return null;
            }
        }

        protected bool SetValueToRfc(string strName, dynamic Value, string Type, int index = -1)
        {
            if (Type.Equals("String"))
            {
                ActionString(strName, Convert.ToString(Value), index);
            }
            else if (Type.Equals("Char"))
            {
                ActionString(strName, Convert.ToString(Value), index);
            }
            else if (Type.Equals("Int32"))
            {
                ActionInt(strName, Convert.ToInt32(Value), index);
            }
            else if (Type.Equals("Int64"))
            {
                long arg = Convert.ToInt64(Value);
                Actionlong(strName, arg, index);
            }
            else if (Type.Equals("Decimal"))
            {
                ActionDecimal(strName, Convert.ToDecimal(Value), index);
            }
            else if (Type.Equals("Double"))
            {
                ActionDouble(strName, Convert.ToDouble(Value), index);
            }
            else if (Type.Equals("Single"))
            {
                float arg2 = Convert.ToSingle(Value);
                ActionSingle(strName, arg2, index);
            }
            else if (!Type.Equals("DateTime"))
            {
                ActionString(strName, Convert.ToString(Value), index);
            }
            else
            {
                DateTime dateTime = Convert.ToDateTime(Value);
                ActionString(strName, dateTime.Date.ToString("yyyyMMdd"), index);
            }
            return true;
        }

        public void RecordNo(string Tag)
        {
            WR.RecordNo(Tag);
        }

        public bool TestFunction()
        {
            WR.WriteToLogFile("Begin Test Function");
            try
            {
                function.Invoke(rfcDest);
                IRfcStructure structure = function.GetStructure("PI_PO_DATA");
                structure.SetValue("EBELN", "4500000000");
                WR.WriteToLogFile("Execute!!!");
                function.SetValue("PI_PO_DATA", structure);
                function.Invoke(rfcDest);
                if (function.GetValue("PE_SUBRC").Equals("0"))
                {
                    IRfcTable table = function.GetTable("PT_DATA");
                    if (table != null)
                    {
                        WR.WriteToLogFile("PT_DATA:" + table.RowCount);
                    }
                    IRfcTable table2 = function.GetTable("PT_RETURN");
                    if (table2 != null)
                    {
                        WR.WriteToLogFile("PT_RETURN is null" + table2.RowCount);
                    }
                    WR.WriteToLogFile("Test OK!");
                }
                WR.WriteToLogFile("Begin Test End");
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***Test Function Fail!");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }


    }


    public class RFC_ToolStructure : RFC_Tool, IDisposable
    {
        private IRfcStructure stru;

        private Action ActionListToRfcTable;

        private bool disposed = false;

        private SafeHandle handle = new SafeFileHandle(IntPtr.Zero, ownsHandle: true);

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    handle.Dispose();
                    stru = null;
                }
                disposed = true;
                base.Dispose(disposing);
            }
        }

        private void SetValueToRfcTable(object value)
        {
            PropertyInfo[] properties = value.GetType().GetProperties();
            PropertyInfo[] array = properties;
            foreach (PropertyInfo propertyInfo in array)
            {
                SetValueToRfc(propertyInfo.Name, propertyInfo.GetValue(value, null), propertyInfo.PropertyType.Name);
            }
        }

        private void GetListPropetyInfoAndSet<T>(List<T> value)
        {
            if (value.Count() > 0)
            {
                SetValueToRfcTable(value[0]);
            }
        }

        public bool SetInputParameter<T>(List<T> value, string StructureName)
        {
            try
            {
                ActionListToRfcTable = delegate
                {
                    function.Invoke(rfcDest);
                    stru = function.GetStructure(StructureName);
                    GetListPropetyInfoAndSet(value);
                    function.SetValue(StructureName, stru);
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool SetInputParameter<T1, T2>(List<T1> value1, string StructureName1, List<T2> value2, string StructureName2)
        {
            try
            {
                ActionListToRfcTable = delegate
                {
                    function.Invoke(rfcDest);
                    stru = function.GetStructure(StructureName1);
                    GetListPropetyInfoAndSet(value1);
                    function.SetValue(StructureName1, stru);
                    stru = null;
                    stru = function.GetStructure(StructureName2);
                    GetListPropetyInfoAndSet(value2);
                    function.SetValue(StructureName2, stru);
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool SetInputParameter<T1, T2, T3>(List<T1> value1, string StructureName1, List<T2> value2, string StructureName2, List<T3> value3, string StructureName3)
        {
            try
            {
                ActionListToRfcTable = delegate
                {
                    function.Invoke(rfcDest);
                    stru = function.GetStructure(StructureName1);
                    GetListPropetyInfoAndSet(value1);
                    function.SetValue(StructureName1, stru);
                    stru = null;
                    stru = function.GetStructure(StructureName2);
                    GetListPropetyInfoAndSet(value2);
                    function.SetValue(StructureName2, stru);
                    stru = null;
                    stru = function.GetStructure(StructureName3);
                    GetListPropetyInfoAndSet(value3);
                    function.SetValue(StructureName3, stru);
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public RFC_ToolStructure(string FunctionName, out bool Sucess)
        {
            WR.WriteDebugLog("***RFC RFC_ToolStructure Create!");
            Sucess = RFC_Init(FunctionName);
            WR.WriteDebugLog("***RFC RFC_ToolStructure Down!");
        }

        protected override bool ReadInputPar()
        {
            try
            {
                ActionString = delegate (string strName, string Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                ActionInt = delegate (string strName, int Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                Actionlong = delegate (string strName, long Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                ActionChar = delegate (string strName, char Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                ActionDecimal = delegate (string strName, decimal Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                ActionDouble = delegate (string strName, double Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                ActionSingle = delegate (string strName, float Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                ActionDateTime = delegate (string strName, DateTime Value, int index)
                {
                    stru.SetValue(strName, Value);
                };
                ActionListToRfcTable();
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***RFC ReadInputPar(RFC_ToolStructure) Fail");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }
    }

    public class RFC_ToolTableStructureMix : RFC_Tool, IDisposable
    {
        private IRfcTable itable = null;

        private IRfcStructure stru = null;

        private bool IsTable = false;

        private Action ActionListToRfcTableStructureMix;

        private bool disposed = false;

        private SafeHandle handle = new SafeFileHandle(IntPtr.Zero, ownsHandle: true);

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    handle.Dispose();
                    itable = null;
                    stru = null;
                }
                disposed = true;
                base.Dispose(disposing);
            }
        }

        private void SetValueToRfcTable(object value, int index)
        {
            PropertyInfo[] properties = value.GetType().GetProperties();
            PropertyInfo[] array = properties;
            foreach (PropertyInfo propertyInfo in array)
            {
                SetValueToRfc(propertyInfo.Name, propertyInfo.GetValue(value, null), propertyInfo.PropertyType.Name, index);
            }
        }

        private void GetListPropetyInfoAndSet<T>(List<T> value)
        {
            if (IsTable)
            {
                for (int i = 0; i < value.Count(); i++)
                {
                    itable.Append();
                    SetValueToRfcTable(value[i], i);
                }
            }
            else if (value.Count() > 0)
            {
                SetValueToRfcTable(value[0], 0);
            }
        }

        public bool SetInputParameter<T1, T2, T3>(List<T1> value1, string StructerName1, List<T2> value2, string TableName2, List<T3> value3, string TableName3, string value4, string SingleName)
        {
            try
            {
                ActionListToRfcTableStructureMix = delegate
                {
                    IsTable = false;
                    ActionString = delegate (string strName, string Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionInt = delegate (string strName, int Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    Actionlong = delegate (string strName, long Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionChar = delegate (string strName, char Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionDecimal = delegate (string strName, decimal Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionDouble = delegate (string strName, double Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionSingle = delegate (string strName, float Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionDateTime = delegate (string strName, DateTime Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    function.Invoke(rfcDest);
                    stru = function.GetStructure(StructerName1);
                    GetListPropetyInfoAndSet(value1);
                    function.SetValue(StructerName1, stru);
                    IsTable = true;
                    ActionString = delegate (string strName, string Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionInt = delegate (string strName, int Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    Actionlong = delegate (string strName, long Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionChar = delegate (string strName, char Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDecimal = delegate (string strName, decimal Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDouble = delegate (string strName, double Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionSingle = delegate (string strName, float Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDateTime = delegate (string strName, DateTime Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    itable = function.GetTable(TableName2);
                    GetListPropetyInfoAndSet(value2);
                    function.SetValue(TableName2, itable);
                    itable = null;
                    itable = function.GetTable(TableName3);
                    GetListPropetyInfoAndSet(value3);
                    function.SetValue(TableName3, itable);
                    SetInputParameter(SingleName, value4);
                    base.ReadInputPar();
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool SetInputParameter<T1, T2, T3>(List<T1> value1, string StructerName1, List<T2> value2, string TableName2, List<T3> value3, string TableName3)
        {
            try
            {
                ActionListToRfcTableStructureMix = delegate
                {
                    IsTable = false;
                    ActionString = delegate (string strName, string Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionInt = delegate (string strName, int Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    Actionlong = delegate (string strName, long Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionChar = delegate (string strName, char Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionDecimal = delegate (string strName, decimal Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionDouble = delegate (string strName, double Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionSingle = delegate (string strName, float Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    ActionDateTime = delegate (string strName, DateTime Value, int index)
                    {
                        stru.SetValue(strName, Value);
                    };
                    function.Invoke(rfcDest);
                    stru = function.GetStructure(StructerName1);
                    GetListPropetyInfoAndSet(value1);
                    function.SetValue(StructerName1, stru);
                    IsTable = true;
                    ActionString = delegate (string strName, string Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionInt = delegate (string strName, int Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    Actionlong = delegate (string strName, long Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionChar = delegate (string strName, char Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDecimal = delegate (string strName, decimal Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDouble = delegate (string strName, double Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionSingle = delegate (string strName, float Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDateTime = delegate (string strName, DateTime Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    itable = function.GetTable(TableName2);
                    GetListPropetyInfoAndSet(value2);
                    function.SetValue(TableName2, itable);
                    itable = null;
                    itable = function.GetTable(TableName3);
                    GetListPropetyInfoAndSet(value3);
                    function.SetValue(TableName3, itable);
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool SetInputParameter<T1, T>(List<T1> value1, string TableName, T Value2, string SingleName)
        {
            try
            {
                ActionListToRfcTableStructureMix = delegate
                {
                    IsTable = true;
                    ActionString = delegate (string strName, string Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionInt = delegate (string strName, int Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    Actionlong = delegate (string strName, long Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionChar = delegate (string strName, char Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDecimal = delegate (string strName, decimal Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDouble = delegate (string strName, double Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionSingle = delegate (string strName, float Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    ActionDateTime = delegate (string strName, DateTime Value, int index)
                    {
                        itable[index].SetValue(strName, Value);
                    };
                    function.Invoke(rfcDest);
                    itable = function.GetTable(TableName);
                    GetListPropetyInfoAndSet(value1);
                    function.SetValue(TableName, itable);
                    SetInputParameter(SingleName, Value2);
                    base.ReadInputPar();
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public RFC_ToolTableStructureMix(string FunctionName, out bool Sucess)
        {
            WR.WriteDebugLog("***RFC RFC_ToolTableStructureMix Create!");
            Sucess = RFC_Init(FunctionName);
            WR.WriteDebugLog("***RFC RFC_ToolTableStructureMix Down!");
        }

        protected override bool ReadInputPar()
        {
            try
            {
                ActionListToRfcTableStructureMix();
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***RFC ReadInputPar(RFC_ToolTableStructureMix) Fail");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }
    }


    public class RFC_ToolTable : RFC_Tool, IDisposable
    {
        private IRfcTable itable = null;

        private string InputStructerName = string.Empty;

        private Action ActionListToRfcTable;

        private bool disposed = false;

        private SafeHandle handle = new SafeFileHandle(IntPtr.Zero, ownsHandle: true);

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    handle.Dispose();
                    itable = null;
                }
                disposed = true;
                base.Dispose(disposing);
            }
        }

        private void SetValueToRfcTable(object value, int index)
        {
            PropertyInfo[] properties = value.GetType().GetProperties();
            PropertyInfo[] array = properties;
            foreach (PropertyInfo propertyInfo in array)
            {
                SetValueToRfc(propertyInfo.Name, propertyInfo.GetValue(value, null), propertyInfo.PropertyType.Name, index);
            }
        }

        private void GetListPropetyInfoAndSet<T>(List<T> value)
        {
            for (int i = 0; i < value.Count(); i++)
            {
                itable.Append();
                SetValueToRfcTable(value[i], i);
            }
        }

        public bool SetInputParameter<T>(List<T> value, string TableName)
        {
            try
            {
                ActionListToRfcTable = delegate
                {
                    function.Invoke(rfcDest);
                    itable = function.GetTable(TableName);
                    GetListPropetyInfoAndSet(value);
                    function.SetValue(TableName, itable);
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool SetInputParameter<T1, T2>(List<T1> value1, string TableName1, List<T2> value2, string TableName2)
        {
            try
            {
                ActionListToRfcTable = delegate
                {
                    function.Invoke(rfcDest);
                    itable = function.GetTable(TableName1);
                    GetListPropetyInfoAndSet(value1);
                    function.SetValue(TableName1, itable);
                    itable = null;
                    itable = function.GetTable(TableName2);
                    GetListPropetyInfoAndSet(value2);
                    function.SetValue(TableName2, itable);
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public bool SetInputParameter<T1, T2, T3>(List<T1> value1, string TableName1, List<T2> value2, string TableName2, List<T3> value3, string TableName3)
        {
            try
            {
                ActionListToRfcTable = delegate
                {
                    function.Invoke(rfcDest);
                    itable = function.GetTable(TableName1);
                    GetListPropetyInfoAndSet(value1);
                    function.SetValue(TableName1, itable);
                    itable = null;
                    itable = function.GetTable(TableName2);
                    GetListPropetyInfoAndSet(value2);
                    function.SetValue(TableName2, itable);
                    itable = null;
                    itable = function.GetTable(TableName3);
                    GetListPropetyInfoAndSet(value3);
                    function.SetValue(TableName3, itable);
                };
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }

        public RFC_ToolTable(string FunctionName, out bool Sucess)
        {
            WR.WriteDebugLog("***RFC RFC_ToolTable Create!");
            Sucess = RFC_Init(FunctionName);
            WR.WriteDebugLog("***RFC RFC_ToolTable Down!");
        }

        protected override bool ReadInputPar()
        {
            try
            {
                ActionString = delegate (string strName, string Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                ActionInt = delegate (string strName, int Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                Actionlong = delegate (string strName, long Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                ActionChar = delegate (string strName, char Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                ActionDecimal = delegate (string strName, decimal Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                ActionDouble = delegate (string strName, double Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                ActionSingle = delegate (string strName, float Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                ActionDateTime = delegate (string strName, DateTime Value, int index)
                {
                    itable[index].SetValue(strName, Value);
                };
                ActionListToRfcTable();
                return true;
            }
            catch (Exception ex)
            {
                WR.WriteToLogFile("***RFC ReadInputPar(RFC_ToolStructure) Fail");
                WR.WriteToLogFile(ex.ToString());
                return false;
            }
        }
    }

    public class InputParameter : IDisposable
    {
        private bool disposed = false;

        private SafeHandle handle = new SafeFileHandle(IntPtr.Zero, ownsHandle: true);

        public string ParameterName { get; private set; }

        public dynamic Value { get; private set; }

        public Type ValueType { get; private set; }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    handle.Dispose();
                }
                disposed = true;
            }
        }

        public InputParameter(string Name, DateTime value)
        {
            SetParameter(Name, value.Date);
        }

        public InputParameter(string Name, char value)
        {
            SetParameter(Name, value);
        }

        public InputParameter(string Name, string value)
        {
            SetParameter(Name, value);
        }

        public InputParameter(string Name, int value)
        {
            SetParameter(Name, value);
        }

        public InputParameter(string Name, double value)
        {
            SetParameter(Name, value);
        }

        public InputParameter(string Name, long value)
        {
            SetParameter(Name, value);
        }

        public InputParameter(string Name, decimal value)
        {
            SetParameter(Name, value);
        }

        public InputParameter(string Name, float value)
        {
            SetParameter(Name, value);
        }

        private void SetParameter<T>(string Name, T value)
        {
            ParameterName = Name;
            Value = value;
            ValueType = value.GetType();
        }

        public T GetInputValue<T>()
        {
            return (T)Value;
        }
    }

    public class OutParameter : IDisposable
    {
        public enum OutType
        {
            Table,
            Structure,
            Single
        }

        private bool disposed = false;

        private SafeHandle handle = new SafeFileHandle(IntPtr.Zero, ownsHandle: true);

        public string ParameterName { get; private set; }

        public dynamic Value { get; private set; }

        public OutType ParameterType { get; private set; }

        public DataTable DTable { get; private set; }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    handle.Dispose();
                    DTable = null;
                }
                disposed = true;
            }
        }

        public OutParameter(string Name, OutType type)
        {
            ParameterName = Name;
            ParameterType = type;
            Value = string.Empty;
            DTable = new DataTable();
        }

        public void Set_rfcTable(DataTable dt)
        {
            DTable = dt;
        }

        public void Set_Value<T>(T Temp)
        {
            Value = Temp;
        }
    }

    public class WriteToFile
    {
        protected OleDbConnection dataConnection = null;

        protected OleDbCommand mySqlCmd = null;

        protected string Tag;

        private bool ConnectToDB()
        {
            try
            {
                dataConnection = new OleDbConnection();
                mySqlCmd = new OleDbCommand();
                string text = ConfigurationManager.AppSettings["BPMDB"];
                string connectionString = text;
                dataConnection.ConnectionString = connectionString;
                dataConnection.Open();
                mySqlCmd.Connection = dataConnection;
                return true;
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString(), "RfcConnectToDB");
                DisconnectToDB();
                return false;
            }
        }

        public bool DisconnectToDB()
        {
            if (mySqlCmd != null)
            {
                mySqlCmd.Dispose();
            }
            if (dataConnection != null)
            {
                if (dataConnection.State == ConnectionState.Open)
                {
                    dataConnection.Close();
                }
                dataConnection.Dispose();
            }
            return true;
        }

        public WriteToFile()
        {
            ConnectToDB();
        }

        public void WriteToLogFile(string message)
        {
            try
            {
                string format = "insert into call_rfc_log(msg,tag)values(N'{0}',N'{1}')";
                format = string.Format(format, message, Tag);
                mySqlCmd.CommandText = format;
                mySqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                WriteLog(message + "| \r\n |" + ex.ToString(), "RfcSyncService_ex");
                throw;
            }
        }

        public void WriteDebugLog(string message)
        {
            WriteToLogFile(message);
        }

        public static void WriteLog(string msg, string FileName)
        {
            //string text = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            //string path = "C:\\temp\\" + FileName + text + ".txt";
            //string directoryName = Path.GetDirectoryName(path);
            //if (!Directory.Exists(directoryName))
            //{
            //	Directory.CreateDirectory(directoryName);
            //}
            //using StreamWriter streamWriter = File.AppendText(path);
            //streamWriter.WriteLine(msg);
        }

        public void RecordNo(string Tag)
        {
            this.Tag = Tag;
        }
    }

}