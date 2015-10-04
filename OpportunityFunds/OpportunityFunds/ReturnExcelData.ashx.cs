using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Script.Serialization;

namespace OpportunityFunds
{
    public class valuelist
    {
        public string key;
        public string value;
        public string easyusekey;
        public string easyusevalue;
    }
    /// <summary>
    /// Summary description for ReturnExcelData
    /// </summary>
    public class ReturnExcelData : IHttpHandler
    {
        List<valuelist> datalist = new List<valuelist>();
        public void ProcessRequest(HttpContext context)
        {
            var excelApp = new Application();
            bool multiple = Boolean.Parse(context.Request["multiple_item"]);
              excelApp.Workbooks.Open("F:\\Calculator V6_With MCA Refinance.xls", Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing);
            var ws = excelApp.Worksheets;
            excelApp.DisplayAlerts = false;
            var worksheet = (Worksheet)ws.get_Item("Calculator");

            Range startCell = worksheet.Cells[3, 5];
            Range endCell = worksheet.Cells[22, 11];
            Range myRange = worksheet.Range[startCell, endCell];

            var selectedproduct = context.Request["product_type"].ToString();
            var Loan_use = context.Request["Loan_Use"].ToString();
            var new_cash = context.Request["New_Cash"].ToString();
            var down_payment = context.Request["Down_pay"].ToString()+"%";
            var term_mnths = context.Request["Terms"].ToString();
            var singleSelect = context.Request["frequency"].ToString();
            worksheet.Cells[3, 6] = selectedproduct;
            worksheet.Cells[4, 6] = Loan_use;
            worksheet.Cells[6, 6] = new_cash;
            worksheet.Cells[8, 6] = down_payment;
            worksheet.Cells[13, 6] = term_mnths;
            worksheet.Cells[15, 6] = singleSelect;
            if (context.Request["easy_pay1"].ToString() == "true")
            {
                var Low = context.Request["Low"];
                var medium = context.Request["medium"];

                var High = context.Request["High"];
                var ccvolume = context.Request["ccvolume"];
                worksheet.Cells[4, 10] = Low;
                worksheet.Cells[5, 10] = medium;
                worksheet.Cells[6, 10] = High;
                worksheet.Cells[9, 11] = down_payment;
            }
            for (int i = 1; i <= 20; i++)
                   
            {
                valuelist temp = new valuelist();
                temp.key = myRange[i, 1].Text;
                temp.value = myRange[i, 2].Text;
                temp.easyusekey = myRange[i, 5].Text;
                temp.easyusevalue = myRange[i, 6].Text;
                datalist.Add(temp);
            }
            if(multiple)
            {
                selectedproduct = context.Request["product_type1"].ToString();
               Loan_use = context.Request["Loan_Use1"].ToString();
                 new_cash = context.Request["New_Cash1"].ToString();

                 down_payment = context.Request["Down_pay1"].ToString() + "%";
                term_mnths = context.Request["Terms1"].ToString();
               singleSelect = context.Request["frequency1"].ToString();
               var interestrate = context.Request["interest_rate"].ToString(); 
                worksheet.Cells[3, 6] = selectedproduct;              
                worksheet.Cells[4, 6] = Loan_use;
                worksheet.Cells[6, 6] = new_cash;
                worksheet.Cells[8, 6] = down_payment;
                worksheet.Cells[13, 6] = term_mnths;
                worksheet.Cells[14, 6] = interestrate;
                worksheet.Cells[15, 6] = singleSelect;
                if (context.Request["easy_pay2"].ToString() == "true")
                {
                    var Low = context.Request["Low1"];
                    var medium = context.Request["medium1"];

                    var High = context.Request["High1"];
                    var ccvolume = context.Request["ccvolume1"];
                    worksheet.Cells[4, 10] = Low;
                    worksheet.Cells[5, 10] = medium;
                    worksheet.Cells[6, 10] = High;
                    worksheet.Cells[9, 11] = down_payment;
                }
                for (int i = 1; i <= 20; i++)
                {
                    valuelist temp = new valuelist();
                    temp.key = myRange[i, 1].Text;
                    temp.value = myRange[i, 2].Text;
                    datalist.Add(temp);
                }
            }
 

            excelApp.Quit();
            JavaScriptSerializer jss = new JavaScriptSerializer();

            string output = jss.Serialize(datalist);
            context.Response.ContentType = "text/plain";
            context.Response.Write(output);
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