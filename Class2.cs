using System;

public class Class2
{
	private void crystalReportViewer1_Load(object sender, EventArgs e)
	{
		CompanyDetails CD = new CompanyDetails();
		CD.Open(0);
		//MessageBox.Show("1");
		//clsReg clsRg = new clsReg();
		ReportDocument cryRpt = new ReportDocument();
		//MessageBox.Show("2");
		string strReportName = "";

		strReportName = "rptSalesQotation.rpt";
		string strPath = Application.StartupPath + "\\" + strReportName;
		cryRpt.Load(strPath);
		//MessageBox.Show("3");
		DataSet ds = new DataSet();

		//MessageBox.Show("4");
		ds = clsReg.GetData_SP(sDocN);
		DataTable dt = ds.Tables[0];
		//MessageBox.Show(ds.Tables[0].Rows.Count.ToString());
		cryRpt.SetDataSource(dt);
		crystalReportViewer1.ReportSource = cryRpt;
		crystalReportViewer1.Visible = true;
	}

	public static DataSet GetData_SP(string sDocN)
	{
		//MessageBox.Show("Entered!!!");
		string strSQLQry = "spSalesQuotaion '" + sDocN + "'";
		DataSet ds = null;
		OdbcConnection objConn = null;
		OdbcDataAdapter Oda = null;
		FOCUSAPILib.FMiscellaneous Fmis = new FOCUSAPILib.FMiscellaneous();
		int ival;
		string strOut1 = "";
		cd.Open(0);
		clsReg.GetRegValue();
		try
		{
			if (iserver == 1)
			{
				ds = new DataSet();
				Oda = new OdbcDataAdapter();
				string s1 = GetConnection(cd.Code);
				objConn = new OdbcConnection(s1);
				objConn.Open();
				Oda.SelectCommand = new OdbcCommand(strSQLQry, objConn);
				Oda.Fill(ds);
			}
			else
			{
				//MessageBox.Show("Comp code:-"+cd.Code);
				//string Qry = ("CompCode=" + cd.Code + ",").Trim() + strSQLQry;
				//Exec focus50M0..spSalesQuotaion '6-R1-R2'
				string Qry = "Exec Focus5" + cd.Code + ".." + strSQLQry;
				ival = Fmis.RemoteFunctionCall("prjSignMAX.clsMain", "GetServerData_Rpt", Qry, ref strOut1);
				string s1 = strOut1;
				//MessageBox.Show("Final Output :- " + s1);
				ds = ConvertXMLToDataSet(s1);
				//MessageBox.Show("In clsReg :- " + ds.Tables.Count.ToString());
			}
		}
		catch (Exception e)
		{
			//info.ShowUserMessage(e.Message);
		}
		finally
		{
			if (Oda != null)
				Oda.Dispose();
			if (objConn != null)
			{
				if (objConn.State == ConnectionState.Open)
					objConn.Close();
				objConn.Dispose();
			}
		}
		return ds;
	}
}
