using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using CallRfc;
using SAP.Middleware.Connector;

namespace Test.API
{
    /// <summary>
    /// Test 的摘要描述
    /// </summary>
    public class Test : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            bool Flag = false;
            try
            {
                string currency = "TWD";
                //currency = "AAA";
                string FunName = "BAPI_EXCHANGERATE_GETDETAIL";
                DataTable dtRate = new DataTable();
                DataTable dtReturn = new DataTable();

                int n;

                if (!string.IsNullOrEmpty(currency))
                {

                    using (RFC_Tool Test = new RFC_Tool(FunName, out Flag))//BAPI_EXCHANGERATE_GETDETAIL
                    {
                        if (Flag)
                        {
                            Test.SetInputParameter("RATE_TYPE", "M");
                            Test.SetInputParameter("FROM_CURR", currency);
                            Test.SetInputParameter("TO_CURRNCY", "TWD");
                            Test.SetInputParameter("DATE", DateTime.Now.ToString("yyyyMMdd"));
                            Test.SetOutParameter("EXCH_RATE", OutParameter.OutType.Structure);
                            Test.SetOutParameter("RETURN", OutParameter.OutType.Structure);

                            if (Test.AccessViaRFC())
                            {
                                dtReturn = Test.GetReturnData("RETURN");
                                if (!string.IsNullOrEmpty(dtReturn.Rows[0]["TYPE"].ToString()))
                                {
                                    //rv.Attributes.Add("success", false);
                                    //rv.Attributes.Add("errMsg", dtReturn.Rows[0]["MESSAGE"].ToString());
                                }
                                else
                                {
                                    dtRate = Test.GetReturnData("EXCH_RATE");
                                    //rv.Attributes.Add("success", true);
                                    //rv.Attributes.Add("rate", dtRate.Rows[0]["EXCH_RATE"].ToString());
                                }
                            }
                            else
                            {
                                //rv.Attributes.Add("success", false);
                                //rv.Attributes.Add("errMsg", "取得RFC回傳值失敗");
                            }
                        }
                        else
                        {
                            //rv.Attributes.Add("success", false);
                            //rv.Attributes.Add("errMsg", "呼叫RFC失敗");
                        }
                    }

                }
                else
                {
                    //rv.Attributes.Add("success", false);
                    //rv.Attributes.Add("errMsg", "匯率選擇錯誤");
                }
                //context.Response.Write(rv);
            }
            catch (Exception ex)
            {
                //Nlog.WriteExNlog(ex, "RfcCurrencyRate", "ProcessRequest");
                throw;
            }
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

    }
}