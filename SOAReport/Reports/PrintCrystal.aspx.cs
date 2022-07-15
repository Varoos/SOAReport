using CrystalDecisions.CrystalReports.Engine;
using SOAReport.Models;
using SOAReport.Reports;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SOAReport
{
    public partial class PrintCrystal : System.Web.UI.Page
    {
        string errors = "";
        string path = "";
        ReportDocument rd = new ReportDocument();
        string FileName = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    System.Net.WebClient client = new System.Net.WebClient();
                    Byte[] buffer = client.DownloadData(Request.QueryString["reportFile"]);

                    if (buffer != null)
                    {
                        Response.ContentType = "application/pdf";
                        Response.AddHeader("content-length", buffer.Length.ToString());
                        Response.BinaryWrite(buffer);
                    }
                }
                catch(Exception ex)
                {

                }
            }
            //rd.Dispose();
        }

        protected void CrystalReportViewer1_Init(object sender, EventArgs e)
        {

        }
    }
}