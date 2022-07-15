using System;

public static void UpdateSQLQuery(string rptSourceURL, string rptNewURL, string DSN, string Database, string UserId, string password, TableLogOnInfo logOnInfo)
{

	ReportDocument rdReport = new ReportDocument();
	rdReport.Load(rptSourceURL, OpenReportMethod.OpenReportByTempCopy);


	// Data souce connection BEFORE/AFTER update through code.
	DataSourceConnections old_rdDataSrcConn = new DataSourceConnections();
	DataSourceConnections new_rdDataSrcConn = new DataSourceConnections();
	old_rdDataSrcConn = rdReport.DataSourceConnections;


	CrystalDecisions.ReportAppServer.Controllers.DataDefController boDataDefController;
	CrystalDecisions.ReportAppServer.DataDefModel.Database boDatabase;
	CrystalDecisions.ReportAppServer.DataDefModel.CommandTable boCommandTable;


	CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rcDocument = rdReport.ReportClientDocument;


	boDataDefController = rcDocument.DataDefController;
	boDatabase = boDataDefController.Database;


	//===============================
	// Main Report Level
	//===============================
	for (int i = 0; i < rdReport.Database.Tables.Count; ++i)
	{
		rdReport.Database.Tables[i].ApplyLogOnInfo(logOnInfo);
		// rcDocument.VerifyDatabase(); // This update doesn't seem to work.


		new_rdDataSrcConn = rdReport.DataSourceConnections;

		// Fall back method of updating if the above method didn't have the effect that we needed.
		if (new_rdDataSrcConn == old_rdDataSrcConn)
		{
			// MessageBox.Show("Connections are the same.");
			CrystalDecisions.ReportAppServer.DataDefModel.ISCRTable rctTable = rcDocument.DataDefController.Database.Tables[i];
			// The following shows how to update connections
			// https://stackoverflow.com/questions/15161227/dynamically-change-database-type-source-etc-in-crystal-reports-for-visual-stud/17797529#17797529
			if (rctTable.ClassName == "CrystalReports.CommandTable")
			{
				CrystalDecisions.ReportAppServer.DataDefModel.CommandTable tbOldCmd = (CrystalDecisions.ReportAppServer.DataDefModel.CommandTable)rctTable;
				CrystalDecisions.ReportAppServer.DataDefModel.CommandTable tbNewCmd = new CrystalDecisions.ReportAppServer.DataDefModel.CommandTable();
				tbNewCmd.Name = tbOldCmd.Name;
				tbNewCmd.Alias = tbOldCmd.Alias;
				tbNewCmd.CommandText = tbOldCmd.CommandText;
				tbNewCmd.Parameters = tbOldCmd.Parameters;
				tbNewCmd.ConnectionInfo = tbOldCmd.ConnectionInfo.Clone(true);
				CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag pbAttr = tbNewCmd.ConnectionInfo.Attributes;


				// pbAttr["Database DLL"] = "crdb_odbc.dll";
				// pbAttr["QE_DatabaseName"] = "NL67S021OUD";
				//pbAttr["DSN"] = DSN;
				pbAttr["QE_DatabaseType"] = "ODBC (RDO)";
				pbAttr["Database Type"] = "ODBC (RDO)";
				// pbAttr["QE_SQLDB"] = "True";
				// pbAttr["SSO Enabled"] = "False";
				pbAttr["QE_ServerDescription"] = DSN;
				pbAttr["QE_DatabaseName"] = Database;
				pbAttr["SSO Enabled"] = false;


				// set connection string                   
				CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag pbLogOnProp = (CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag)pbAttr["QE_LogonProperties"];
				pbLogOnProp.RemoveAll();
				// strangely comma seperated values instead of semicolon seperated values are needed here
				//pbLogOnProp.FromString("Provider=IBMDA400,Data Source=" + sServerName + ",Initial Catalog=" + sDBName + ",User ID=" + sUserId + ",Password=" + sPwd + ",Convert Date Time To Char=TRUE,Catalog Library List=,Cursor Sensitivity=3");
				//pbLogOnProp.FromString("DSN="+DSN+",Database=" + Database + ",User ID=" + UserId + ",Use DSN Default Properties=False,PreQEServerName="+DSN+"");
				pbLogOnProp.Add("DSN", DSN);
				pbLogOnProp.Add("Database", Database);
				pbLogOnProp.Add("User ID", UserId);
				pbLogOnProp.Add("Use DSN Default Properties", "False");
				pbLogOnProp.Add("PreQEServerName", DSN);
				pbLogOnProp.Add("Database Type", "ODBC (RDO)");
				//PropertyBag connectionAttributes = new PropertyBag();
				//connectionAttributes.Add("Auto Translate", "-1");
				tbNewCmd.ConnectionInfo.UserName = UserId;
				tbNewCmd.ConnectionInfo.Password = password;


				rcDocument.DatabaseController.SetTableLocation(tbOldCmd, tbNewCmd);
				//rcDocument.VerifyDatabase(); // Doesn't work after using this method.
			}
		}
	}





	public ActionResult ExportPrint()
	{
		var customer = "";


		APICalls.logging("Hello1");
		ReportDocument Cr = new ReportDocument();
		APICalls.logging("Hello2");
		SqlConnectionStringBuilder SConn = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["Connection"].ConnectionString);

		string s1FDate = Session["sFDate"].ToString();
		string s1TDate = Session["sTDate"].ToString();
		string s1CustId = Session["sCustId"].ToString();
		string s1user = Session["s1LoginId"].ToString();
		APICalls.logging(s1FDate);
		APICalls.logging(s1TDate);
		APICalls.logging(s1CustId);
		APICalls.logging(s1user);

		APICalls.logging(thisConnectionString);
		SqlConnection con = new SqlConnection(thisConnectionString);
		if (con.State != ConnectionState.Open)
		{
			con.Open();
		}

		ReportDocument rd = new ReportDocument();
		string crydata = "";

		if (s1CustId == "0")
		{
			customer = "%";
		}
		else
		{
			customer = s1CustId;
		}

		crydata = "Exec [dbo].[spSOA] '" + s1FDate + "','" + s1TDate + "','" + customer + "'," + s1user;

		APICalls.logging(crydata);
		//string ccd = System.Configuration.ConfigurationManager.AppSettings["CompCode"].ToString();
		//APICall.logging(ccd);
		DataSet ds = FocusRegistry.GetData(crydata);
		var rptSource = ds.Tables[0];
		string strRptPath = ("C:\\inetpub\\wwwroot\\SOAKasimy\\Rpt\\rptSOA.rpt");
		rd.Load(strRptPath);
		rd.SetParameterValue("@FromDate", s1FDate);
		rd.SetParameterValue("@ToDate", s1TDate);
		rd.SetParameterValue("@Customerid", customer);
		rd.SetParameterValue("@sUser", s1user);

		rd.SetParameterValue("@FromDate", s1FDate, rd.Subreports[0].Name.ToString());
		rd.SetParameterValue("@ToDate", s1TDate, rd.Subreports[0].Name.ToString());
		rd.SetParameterValue("@Customerid", customer, rd.Subreports[0].Name.ToString());
		rd.SetParameterValue("@sUser", s1user, rd.Subreports[0].Name.ToString());


		// Report connection
		ConnectionInfo connInfo = new ConnectionInfo();
		connInfo.ServerName = SConn.DataSource;
		connInfo.DatabaseName = SConn.InitialCatalog;
		connInfo.UserID = SConn.UserID;
		connInfo.Password = SConn.Password;


		TableLogOnInfo tableLogOnInfo = new TableLogOnInfo();
		tableLogOnInfo.ConnectionInfo = connInfo;
		foreach (CrystalDecisions.CrystalReports.Engine.Table table in rd.Database.Tables)
		{
			table.ApplyLogOnInfo(tableLogOnInfo);
			table.LogOnInfo.ConnectionInfo.ServerName = connInfo.ServerName;
			table.LogOnInfo.ConnectionInfo.DatabaseName = connInfo.DatabaseName;
			table.LogOnInfo.ConnectionInfo.UserID = connInfo.UserID;
			table.LogOnInfo.ConnectionInfo.Password = connInfo.Password;

		}
		APICalls.logging("hello");
		Response.Buffer = false;
		Response.ClearContent();
		Response.ClearHeaders();

		//APICall.logging("Export");
		APICalls.logging("hello1");
		Stream stream = rd.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
		stream.Seek(0, SeekOrigin.Begin);
		//CR_reportDocument.ExportToDisk(ExportFormatType.PortableDocFormat, "C:\\Data\\converted2crystal\\EDC_PLANNING_RPT_params.pdf");
		//Stream stream = rd.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Response, true, "SOA");


		return File(stream, "application/pdf", "SOA.pdf");
		//if (rptSource != null && rptSource.GetType().ToString() != "System.String")
		//	rd.SetDataSource(rptSource);
		//rd.ExportToHttpResponse(ExportFormatType.PortableDocFormat, System.Web.HttpContext.Current.Response, false, "SOA");
		//APICalls.logging("Hi");
		//return Json(1, JsonRequestBehavior.AllowGet);

	}
	