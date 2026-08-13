/*
 * Created by SharpDevelop.
 * User: jgazon
 * Date: 25-06-20
 * Time: 16:44
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Text.RegularExpressions;
using System.Linq;
using System.Windows;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Diagnostics;


namespace Dispatch
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			db_read();
			//System.Threading.Thread.Sleep(15000);
			//Launch_Notams_Analysis();
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		public static string IsOpsType(string OpsType, string location)
		{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO='"+location+"'";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
     		string IsOps="";
     		
     		if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(OpsType=="LH")
            		{
            			if(!dBreader.IsDBNull(3)) IsOps = dBreader.GetString(3);
            		}
            		if(OpsType=="FedEx")
            		{
            			if(!dBreader.IsDBNull(4)) IsOps = dBreader.GetString(4);
            		}
            		if(OpsType=="Charters")
            		{
            			if(!dBreader.IsDBNull(5)) IsOps = dBreader.GetString(5);
            		}
            	}
            }
     		conn.Close();

			return IsOps;
		}
		public static string GetIATA(string location)
		{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO='"+location+"'";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
     		string IATA="";
     		if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(2)) IATA = dBreader.GetString(2);
            	}
     		}
     		conn.Close();
     		return IATA;
		}

		void Btn_addTAFClick(object sender, EventArgs e)
		{
			string pattern = @"\b[A-Z]{4,4}\b [0-9]{6}Z";
      		Regex rgx = new Regex(pattern);
      		string pastedTAF = RchTxt_TAF.Text;
      		pastedTAF = Regex.Replace(pastedTAF, @"\t|\n|\r", " ");
      		pastedTAF = Regex.Replace(pastedTAF, @"\s+", " ");
      		string[] stationsTAF = Regex.Split(pastedTAF, pattern);
      		
      		int i=0;
      		int stationsTAFlength = stationsTAF.Length;
      		string result = stationsTAFlength.ToString()+"<br />";
  			foreach (string value in stationsTAF)
  			{
				stationsTAF[i] = value;
				int TAFlength = stationsTAF[i].Length;
				string TAFend = "";
				if(TAFlength>3)
				{
					TAFend = stationsTAF[i].Substring(TAFlength-4, 4);
					if(TAFend=="TAF ") stationsTAF[i] = stationsTAF[i].Substring(0,TAFlength-4);
				}
				TAFlength = stationsTAF[i].Length;
				if(TAFlength>8)
				{
					TAFend= stationsTAF[i].Substring(TAFlength-8, 8);
					//stationsTAF[i]= TAFend;
					if(TAFend=="TAF COR " |TAFend=="TAF AMD ") stationsTAF[i] = stationsTAF[i].Substring(0,TAFlength-8);
				}
				
  				i++;
  			}
      		
      		string[] stationICAO= new string[i];
      		string[] stationTime= new string[i];
      		i=1;
     		foreach (Match match in rgx.Matches(pastedTAF))
     		{
     			stationICAO[i]=match.Value;
     			stationTime[i]=stationICAO[i].Substring(5,6);
     			stationICAO[i]=stationICAO[i].Substring(0,4);
     			i++;
     		}
     		
     		
     		System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
			string deletelog = "DELETE * FROM TAF_analysis";

			OleDbCommand commandedelete = new OleDbCommand(deletelog, conn);
			// Execution
				
			commandedelete.ExecuteNonQuery();
			conn.Close();
			
			conn.Open();
			
			for (int j=1;j<i;j++)
			{	
				string insertlog = "INSERT INTO TAF_analysis (ICAO,TimeIssued) VALUES ('"+stationICAO[j]+"','"+stationTime[j]+"')";
				OleDbCommand commandeinsert = new OleDbCommand(insertlog, conn);
				commandeinsert.ExecuteNonQuery();
			}
				
			conn.Close();
			i=0;
     		//result +=stationsTAF[5]+"<br />";
     		result += "<table>";

  			foreach (string value in stationsTAF)
  			{
  				//string patternTrend = @"((<=|BECMG |PROB40 |TEMPO |PROB30 |FM)[0-9]{4})";
  				string patternTrend = @"(<=|BECMG [0-9]{4}|TEMPO [0-9]{4}|PROB30 [0-9]{4}|PROB40 [0-9]{4}|FM[0-9]{4})";
      			Regex rgxTrend = new Regex(patternTrend);
      		
      			string[] stationsTAFsplit = Regex.Split(value, patternTrend);
      			int iSplit =0;
      			int sizeSplit = stationsTAFsplit.Length;
      			
      			foreach (string valueSplit in stationsTAFsplit) 
      			{
      				int splitLength = valueSplit.Length;
      				string splitEnd="";
      				if (splitLength>7) splitEnd = valueSplit.Substring(splitLength-7,7);
      				//testProb +=">>>"+splitEnd+"<<<";
      				//if (splitEnd=="PROB30 " && iSplit< sizeSplit)
      				if (splitEnd=="PROB30 ")
      				{
      					stationsTAFsplit[iSplit]= stationsTAFsplit[iSplit].Substring(0,splitLength-7);
      					stationsTAFsplit[iSplit+1]= "PROB30 "+ stationsTAFsplit[iSplit+1];
      				}
      				if (splitEnd=="PROB40 ")
      				{
      					stationsTAFsplit[iSplit]= stationsTAFsplit[iSplit].Substring(0,splitLength-7);
      					stationsTAFsplit[iSplit+1]= "PROB40 "+ stationsTAFsplit[iSplit+1];
      				}
      				//MEME Operation mais 10 iso 7 su a retour ligne 
      				if (splitLength>10) splitEnd = valueSplit.Substring(splitLength-10,7);
      				//testProb +=">>>"+splitEnd+"<<<";
      				//if (splitEnd=="PROB30 " && iSplit< sizeSplit)
      				if (splitEnd=="PROB30 ")
      				{
      					stationsTAFsplit[iSplit]= stationsTAFsplit[iSplit].Substring(0,splitLength-10);
      					stationsTAFsplit[iSplit+1]= "PROB30 "+ stationsTAFsplit[iSplit+1];
      				}
      				if (splitEnd=="PROB40 ")
      				{
      					stationsTAFsplit[iSplit]= stationsTAFsplit[iSplit].Substring(0,splitLength-10);
      					stationsTAFsplit[iSplit+1]= "PROB40 "+ stationsTAFsplit[iSplit+1];
      				}
      				
      				iSplit ++;
      			}
      			
      			string stationSplittedBr="";
      			
      			foreach (string valueSplit in stationsTAFsplit) 
      			{
      				
      				string transValue ="-";
      				if(valueSplit.Length>0) transValue = valueSplit.Substring(0,1);

      				
      				if(transValue=="/" || transValue=="0")
      				{
      					stationSplittedBr += valueSplit;
      				}
      				else
      				{
      					stationSplittedBr += "<br />"+valueSplit;
      				}
      			}
      			
      			string patternBr = @"<br />";
      			Regex rgxBr = new Regex(patternBr);
      		
      			string[] TAFbrSplit = Regex.Split(stationSplittedBr, patternBr);
      			string ceilVis="";
      			bool trend = false;
      			bool visSelected =false;
      			bool ceilCatI = false;
      			bool ceilTresh = false;
      			bool visCatI = false;
      			bool visTresh = false;
      			bool windTrend=false;
      			string testWind="";
				string TS="";
				string SN="";
      			
      			foreach (string valueBr in TAFbrSplit) 
      			{
      				//VIS & CEILING
      				string patternVis = @"(?<= )\b[0-9]{4,4}\b(?= )";
      				Regex rgxVis = new Regex(patternVis);
      				string twoFirst="";
      				if(valueBr.Length>=3) twoFirst=valueBr.Substring(0,2);
      				foreach (Match matchVis in rgxVis.Matches(valueBr))
  					{      					   				
      					string stringVis=matchVis.ToString();
      					int intVis= int.Parse(stringVis);
      					
      					if(intVis<550)
      					{
      						visCatI=true;
      						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit)) trend=true;
       					}
      					else if(intVis<=1000 && intVis>=550)
      					{
      						visTresh= true;
      						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit))trend=true;
       					} 
  					}
      				string patternCeil = @"(?<=BKN|OVC|VV)[0-9]{3}";
      				Regex rgxCeil = new Regex(patternCeil);
      				foreach (Match matchCeil in rgxCeil.Matches(valueBr))
  					{
      					string stringCeil = matchCeil.ToString();
      					int intCeil = int.Parse(stringCeil);
      					
      					if(intCeil<=2)
      					{
      						ceilCatI=true;
      						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit)) trend=true;
      					}
      					else if(intCeil<=4 && intCeil>2)
      					{
      						ceilTresh=true;
      						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit))trend=true;
      					}
      				}
      				
      				if(ceilCatI||visCatI)
      				{
      					ceilVis+="<span style=\"color:red;\">"+valueBr+"</span>";     					
      				}
      				else if((ceilTresh||visTresh) && !ceilCatI && !visCatI)
      				{
      					ceilVis+=valueBr;
      				}
      				else if(trend && (twoFirst=="BE"||twoFirst=="FM"))
      				{
      					ceilVis+="<span style=\"color:blue;\">"+valueBr+"</span>";
      					trend=false;
      				}
      				ceilCatI=false;
      				ceilTresh=false;
      				visCatI=false;
      				visTresh=false;
      				
      				//WIND
      				string patternWind = @"([a-zA-Z0-9]{5,8})KT";
      				Regex rgxWind = new Regex(patternWind);
      				twoFirst="";
      				if(valueBr.Length>=3) twoFirst=valueBr.Substring(0,2);
      				      				
      				foreach (Match matchWind in rgxWind.Matches(valueBr))
  					{
      					
      					string stringWind=matchWind.ToString();
      					
      					string fourFive =stringWind.Substring(3,2);
      					//check for Gusts
      					if(stringWind.Length==10)fourFive=stringWind.Substring(6,2);
      					int intWind= int.Parse(fourFive);
      					
      					if(intWind>44)
      					{
      						testWind+="<span style=\"color:red;\">"+valueBr+"</span>";
       						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit))windTrend=true;
       					}
      					else if(intWind>34 && intWind<45) 
      					{
      						testWind+=valueBr;
      						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit))windTrend=true;
       					}
      					else if(intWind<34 && windTrend==true && (twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit)))
      					{
      						testWind+="<span style=\"color:blue;\">"+valueBr+"</span>";
      						windTrend=false;
      					}
  					}
      				//WIND MPS
      				string patternWindMPS = @"([a-zA-Z0-9]{5,8})MPS";
      				Regex rgxWindMPS = new Regex(patternWindMPS);
      				twoFirst="";
      				if(valueBr.Length>=3) twoFirst=valueBr.Substring(0,2);
      				      				
      				foreach (Match matchWindMPS in rgxWindMPS.Matches(valueBr))
  					{
      					
      					string stringWindMPS=matchWindMPS.ToString();
      					string fourFive =stringWindMPS.Substring(3,2);
      					
      					//check for Gusts
      					if(stringWindMPS.Length==11)fourFive=stringWindMPS.Substring(6,2);
      					int intWindMPS= int.Parse(fourFive);
      					
      					if(intWindMPS>22)
      					{
      						//Web_TAF.DocumentText = valueBr;
      						testWind+="<span style=\"color:red;\">"+valueBr+"</span>";
       						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit))windTrend=true;
       					}
      					else if(intWindMPS>17 && intWindMPS<22) 
      					{
      						testWind+=valueBr;
      						if(twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit))windTrend=true;
       					}
      					else if(intWindMPS<34 && windTrend==true && (twoFirst=="BE"||twoFirst=="FM"||twoFirst.All(char.IsDigit)))
      					{
      						testWind+="<span style=\"color:blue;\">"+valueBr+"</span>";
      						windTrend=false;
      					}
  					}
      				
      				//TSRA
      				string patternTS = "TS";
      				Regex rgxTS = new Regex(patternTS);
      				
      				foreach (Match matchTS in rgxTS.Matches(valueBr))
  					{
      					TS+=valueBr;
      				}
      				
      				//Snow
      				string patternSN = "SN|FZRA|FZDZ";
      				Regex rgxSN = new Regex(patternSN);
      				
      				foreach (Match matchSN in rgxSN.Matches(valueBr))
  					{
      					SN+=valueBr;
      				}
       			}
      			
      			conn.Open();
      			string updatelog = "UPDATE TAF_analysis SET Vis_Ceiling='"+ceilVis+"',Wind='"+testWind+"',TS='"+TS+"',Snow='"+SN+"' WHERE ICAO='"+stationICAO[i]+"'";
				OleDbCommand commandeupdate = new OleDbCommand(updatelog, conn);
				commandeupdate.ExecuteNonQuery();
      			conn.Close();
      			
  				result += "<tr><td>"+stationICAO[i]+"</td><td>" + testWind +"</td></tr>";

  				i++;
  				
  			}
  			result+="</table>";
  			
  				//Web_TAF.DocumentText=result;
  			
  			//date & time of analysis
  		string timeTAF = "";
			string time = DateTime.Now.ToString("HH:mm");
			string today = DateTime.Now.ToShortDateString();
			timeTAF = time +"(CET) "+ today;
  			conn.Open();
      		string updateTimeTAF = "UPDATE TAF_timestamp SET TimeTAF='"+timeTAF+"' WHERE ID=1";
			OleDbCommand commandeupdateTimeTAF = new OleDbCommand(updateTimeTAF, conn);
			commandeupdateTimeTAF.ExecuteNonQuery();
      		conn.Close();
  			
     		db_read();
     		
		}
		public void db_read()
		{			
			APT_List.AutoScroll=true;
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			//Header
			conn.Open();
			
			string timeTAF="";
			var queryTimeTAF = "SELECT * FROM TAF_timestamp WHERE ID=1";
     		OleDbCommand commandTimeTAF = new OleDbCommand(queryTimeTAF, conn);
     		OleDbDataReader readerTimeTAF = commandTimeTAF.ExecuteReader();
     		while (readerTimeTAF.Read())
     		{
     			if(readerTimeTAF.GetString(1)!="") timeTAF=readerTimeTAF.GetString(1);
     		}
			
			string report ="";
			report+="For new TAF analysis, paste the APT List in TAF's on <a href=\"https://aviationweather.gov/taf\">https://aviationweather.gov/taf</a><br />" +
				"Last analysis : "+timeTAF+"<hr />";
			
			conn.Close();
			conn.Open();
			
			report+="<table style=\"text-align: left; font-size:12px\">";
			
			//Long Haul
			report+="<tr><th colspan=\"5\"><span style=\"font-weight:bold; font-size:16px; color : RoyalBlue;\">" +
				"Long Haul Ops :</span></th></tr><tr><th colspan=\"5\">"+
				"<tr><th colspan=\"5\"><b>Ceiling & Visibility</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";
			
			//LH Vis& Ceiling
			var queryLHVisCeiling = "SELECT ICAO,Vis_Ceiling FROM TAF_analysis WHERE Vis_Ceiling IS NOT NULL";
     		OleDbCommand commandLHVisCeiling = new OleDbCommand(queryLHVisCeiling, conn);
     		OleDbDataReader readerLHVisCeiling = commandLHVisCeiling.ExecuteReader();
     		
     		bool TAFresult=false;
			while (readerLHVisCeiling.Read())
        	{
				string ICAO = readerLHVisCeiling.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string LH="";
     			while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(3)!="")LH=readerIATA.GetString(3);

     			}
     			
				string Vis_Ceiling = readerLHVisCeiling.GetString(1);
				if(Vis_Ceiling!="" && LH=="Yes")
       			{ 						
					report += "<b style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerLHVisCeiling.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//LH WIND
			report+="</th></tr><tr><th colspan=\"5\"><b>Wind</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";
			var queryLHWind = "SELECT ICAO,Wind FROM TAF_analysis WHERE Wind IS NOT NULL";
			
    		OleDbCommand commandLHWind = new OleDbCommand(queryLHWind, conn);
    		OleDbDataReader readerLHWind = commandLHWind.ExecuteReader();

			while (readerLHWind.Read())
    		{
				string ICAO = readerLHWind.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string LH="";
     			while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(3)!="")LH=readerIATA.GetString(3);

     			}
				string Wind = readerLHWind.GetString(1);
				if(Wind!="" && LH=="Yes")
    		   	{
					report+= "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerLHWind.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//LH TS
			report+="</th></td><tr><th colspan=\"5\"><b>Thunderstorms</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";	
			var queryLHTS = "SELECT ICAO,TS FROM TAF_analysis WHERE TS IS NOT NULL";
			
     		OleDbCommand commandLHTS = new OleDbCommand(queryLHTS, conn);
     		OleDbDataReader readerLHTS = commandLHTS.ExecuteReader();

			while (readerLHTS.Read())
        	{
				string ICAO = readerLHTS.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string LH="";
				while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(3)!="")LH=readerIATA.GetString(3);

     			}
				string TS = readerLHTS.GetString(1);
				if(TS!="" && LH=="Yes")
       			{
					report += "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerLHTS.GetString(1)+"<br />";
						TAFresult=true;
					}
				}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//Snow
			report+="</th></td><tr><th colspan=\"5\"><b>Snow</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";	
			var queryLHSN = "SELECT ICAO,Snow FROM TAF_analysis WHERE Snow IS NOT NULL";
			
     		OleDbCommand commandLHSN = new OleDbCommand(queryLHSN, conn);
     		OleDbDataReader readerLHSN = commandLHSN.ExecuteReader();

			while (readerLHSN.Read())
        	{
				string ICAO = readerLHSN.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string LH="";
				while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(3)!="")LH=readerIATA.GetString(3);

     			}
				string SN = readerLHSN.GetString(1);
				if(SN!="" && LH=="Yes")
       			{
					report += "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerLHSN.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span></th></tr>";
			report+="<tr><th colspan=\"5\"><hr /></th></tr>";
			
			conn.Close();
			conn.Open();
			
			//FedEx
			report+="<tr><th colspan=\"5\"><span style=\"font-weight:bold; font-size:16px; color : DarkMagenta;\">FedEx Ops :</span></th></tr><tr><th colspan=\"5\">"+
				"<tr><th colspan=\"5\"><b>Ceiling & Visibility</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";
			
			//FEDEX Vis& Ceiling
			var queryVisCeiling = "SELECT ICAO,Vis_Ceiling FROM TAF_analysis WHERE Vis_Ceiling IS NOT NULL";
     		OleDbCommand commandVisCeiling = new OleDbCommand(queryVisCeiling, conn);
     		OleDbDataReader readerVisCeiling = commandVisCeiling.ExecuteReader();
     		TAFresult=false;
			while (readerVisCeiling.Read())
        	{
				string ICAO = readerVisCeiling.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string FedEx="";
     			while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(4)!="")FedEx=readerIATA.GetString(4);

     			}
     			
				string Vis_Ceiling = readerVisCeiling.GetString(1);
				if(Vis_Ceiling!="" && FedEx=="Yes")
       			{ 						
					report += "<b style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerVisCeiling.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//FEDEX WIND
			report+="</th></tr><tr><th colspan=\"5\"><b>Wind</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";
			var queryWind = "SELECT ICAO,Wind FROM TAF_analysis WHERE Wind IS NOT NULL";
			
    		OleDbCommand commandWind = new OleDbCommand(queryWind, conn);
    		OleDbDataReader readerWind = commandWind.ExecuteReader();

			while (readerWind.Read())
    		{
				string ICAO = readerWind.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string FedEx="";
     			while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(4)!="")FedEx=readerIATA.GetString(4);

     			}
				string Wind = readerWind.GetString(1);
				if(Wind!="" && FedEx=="Yes")
    		   	{
					report+= "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerWind.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//FedEx TS
			report+="</th></td><tr><th colspan=\"5\"><b>Thunderstorms</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";	
			var queryTS = "SELECT ICAO,TS FROM TAF_analysis WHERE TS IS NOT NULL";
			
     		OleDbCommand commandTS = new OleDbCommand(queryTS, conn);
     		OleDbDataReader readerTS = commandTS.ExecuteReader();

			while (readerTS.Read())
        	{
				string ICAO = readerTS.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string FedEx="";
				while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(4)!="")FedEx=readerIATA.GetString(4);

     			}
				string TS = readerTS.GetString(1);
				if(TS!="" && FedEx=="Yes")
       			{
					report += "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerTS.GetString(1)+"<br />";
						TAFresult=true;
					}
				}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//FedEx Snow
			report+="</th></td><tr><th colspan=\"5\"><b>Snow</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";	
			var querySN = "SELECT ICAO,Snow FROM TAF_analysis WHERE Snow IS NOT NULL";
			
     		OleDbCommand commandSN = new OleDbCommand(querySN, conn);
     		OleDbDataReader readerSN = commandSN.ExecuteReader();

			while (readerSN.Read())
        	{
				string ICAO = readerSN.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string FedEx="";
				while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(4)!="")FedEx=readerIATA.GetString(4);

     			}
				string SN = readerSN.GetString(1);
				if(SN!="" && FedEx=="Yes")
       			{
					report += "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerSN.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span></th></tr>";
			report+="<tr><th colspan=\"5\"><hr /></th></tr>";
			
     		conn.Close();
     		conn.Open();
     		
     		//Charters
			report+="<tr><th colspan=\"5\"><span style=\"font-weight:bold; font-size:16px; color : Green;\">" +
				"Charters Ops :</span></th></tr><tr><th colspan=\"5\">"+
				"<tr><th colspan=\"5\"><b>Ceiling & Visibility</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";
			
			//Charters Vis& Ceiling
			var queryChartersVisCeiling = "SELECT ICAO,Vis_Ceiling FROM TAF_analysis WHERE Vis_Ceiling IS NOT NULL";
     		OleDbCommand commandChartersVisCeiling = new OleDbCommand(queryChartersVisCeiling, conn);
     		OleDbDataReader readerChartersVisCeiling = commandChartersVisCeiling.ExecuteReader();
     		
     		TAFresult=false;
			while (readerChartersVisCeiling.Read())
        	{
				string ICAO = readerChartersVisCeiling.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string Charters="";
     			while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(5)!="")Charters=readerIATA.GetString(5);

     			}
     			
				string Vis_Ceiling = readerChartersVisCeiling.GetString(1);
				if(Vis_Ceiling!="" && Charters=="Yes")
       			{ 						
					report += "<b style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerChartersVisCeiling.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//Charters WIND
			report+="</th></tr><tr><th colspan=\"5\"><b>Wind</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";
			var queryChartersWind = "SELECT ICAO,Wind FROM TAF_analysis WHERE Wind IS NOT NULL";
			
    		OleDbCommand commandChartersWind = new OleDbCommand(queryChartersWind, conn);
    		OleDbDataReader readerChartersWind = commandChartersWind.ExecuteReader();

			while (readerChartersWind.Read())
    		{
				string ICAO = readerChartersWind.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string Charters="";
     			while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(5)!="")Charters=readerIATA.GetString(5);

     			}
				string Wind = readerChartersWind.GetString(1);
				if(Wind!="" && Charters=="Yes")
    		   	{
					report+= "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerChartersWind.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;
			report+="</span>";
			
			conn.Close();
			conn.Open();
			
			//Charters TS
			report+="</th></td><tr><th colspan=\"5\"><b>Thunderstorms</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";	
			var queryChartersTS = "SELECT ICAO,TS FROM TAF_analysis WHERE TS IS NOT NULL";
			
     		OleDbCommand commandChartersTS = new OleDbCommand(queryChartersTS, conn);
     		OleDbDataReader readerChartersTS = commandChartersTS.ExecuteReader();

			while (readerChartersTS.Read())
        	{
				string ICAO = readerChartersTS.GetString(0);
				var queryIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandIATA = new OleDbCommand(queryIATA, conn);
     			OleDbDataReader readerIATA = commandIATA.ExecuteReader();
				string IATA="XXX";
				string Charters="";
				while (readerIATA.Read()) 
     			{
     					if (readerIATA.GetString(2)!="")IATA=readerIATA.GetString(2);
     					if (readerIATA.GetString(5)!="")Charters=readerIATA.GetString(5);

     			}
				string TS = readerChartersTS.GetString(1);
				if(TS!="" && Charters=="Yes")
       			{
					report += "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerChartersTS.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;

			
			conn.Close();
			conn.Open();

            report+="</span>";
			
            conn.Close();
			conn.Open();
            
            //Charters Snow
			report+="</th></td><tr><th colspan=\"5\"><b>Snow</b></th></tr><tr><th colspan=\"5\">";
			report+="<span style =\"font-family: Courier New; font-weight:normal;\">";	
			var queryChartersSN = "SELECT ICAO,Snow FROM TAF_analysis WHERE Snow IS NOT NULL";
			
     		OleDbCommand commandChartersSN = new OleDbCommand(queryChartersSN, conn);
     		OleDbDataReader readerChartersSN = commandChartersSN.ExecuteReader();

			while (readerChartersSN.Read())
        	{
				string ICAO = readerChartersSN.GetString(0);
				var queryChartersIATA = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=\'"+ICAO+"\'";
				OleDbCommand commandChartersIATA = new OleDbCommand(queryChartersIATA, conn);
     			OleDbDataReader readerChartersIATA = commandChartersIATA.ExecuteReader();
				string IATA="XXX";
				string Charters="";
				while (readerChartersIATA.Read()) 
     			{
     					if (readerChartersIATA.GetString(2)!="")IATA=readerChartersIATA.GetString(2);
     					if (readerChartersIATA.GetString(5)!="")Charters=readerChartersIATA.GetString(5);

     			}
				string SN = readerChartersSN.GetString(1);
				if(SN!="" && Charters=="Yes")
       			{
					report += "<b  style=\"color : DarkBlue;font : bold;\">"+ICAO+" - "+IATA+" : </b>"+readerChartersSN.GetString(1)+"<br />";
					TAFresult=true;
				}
			}
			if(!TAFresult)report+= "Nil";
			TAFresult=false;

			report+="</span></th></tr></table>";
     		conn.Close();
     		
     		Web_TAF.DocumentText = report;
					
		
       		List<Label> itemsToRemove = new List<Label>();
			foreach (Label label in Tabs.SelectedTab.Controls.OfType<Label>())
			{
    			if (label.Tag != null && label.Tag.ToString() == "dispose")
    			{
      		  		itemsToRemove.Add(label);
    			}
			}
			foreach (Label label in itemsToRemove)
			{
    			Controls.Remove(label);
    			label.Dispose();
			}
			
			List<RichTextBox> rchtxtboxToRemove = new List<RichTextBox>();
			foreach (RichTextBox rchtxtbox in Tabs.SelectedTab.Controls.OfType<RichTextBox>())
			{
    			if (rchtxtbox.Tag != null && rchtxtbox.Tag.ToString() == "dispose")
    			{
      		  		rchtxtboxToRemove.Add(rchtxtbox);
    			}
			}
			foreach (RichTextBox rchtxtbox in rchtxtboxToRemove)
			{
    			Controls.Remove(rchtxtbox);
    			rchtxtbox.Dispose();
			}
			
			List<CheckBox> chckboxToRemove = new List<CheckBox>();
			foreach (CheckBox chckbox in Tabs.SelectedTab.Controls.OfType<CheckBox>())
			{
    			if (chckbox.Tag != null && chckbox.Tag.ToString() == "dispose")
    			{
      		  		chckboxToRemove.Add(chckbox);
    			}
			}
			
			foreach (CheckBox chckbox in chckboxToRemove)
			{
    			Controls.Remove(chckbox);
    			chckbox.Dispose();
			}
			
			List<Button> buttonsToRemove = new List<Button>();
			foreach (Button button in Tabs.SelectedTab.Controls.OfType<Button>())
			{
    			if (button.Tag != null && button.Tag.ToString() == "dispose")
    			{
      		  		buttonsToRemove.Add(button);
    			}
			}
			foreach (Button button in buttonsToRemove)
			{
    			Controls.Remove(button);
    			button.Dispose();
			}

// APT List
			if(btn_editList.Text=="Close List")
			{
				conn.Open();
	       		var queryAPTList = "Select * From Stations_ICAO_IATA ORDER BY ICAO";
	     		OleDbCommand commandAPTList = new OleDbCommand(queryAPTList, conn);
	     		OleDbDataReader readerAPTList = commandAPTList.ExecuteReader();
	     		
	     		int int_APT_ID = 0;
	     		string string_APT_ICAO ="";
	     		string string_APT_IATA ="";
	     		string string_APT_LH="";
	     		string string_APT_FedEx="";
	     		string string_APT_Charters ="";
	     		
	     		TxtBox_APT_ICAO.Text="";
	     		TxtBox_APT_IATA.Text="";
	     		ChckBx_APT_LH.Checked=false;
	     		ChckBx_APT_FedEx.Checked=false;
	     		ChckBx_APT_Charters.Checked=false;
	     		
	     		Button[] del_APT_Buttons = new Button[300];
	     		Button[] edit_APT_Buttons = new Button[300];
	     		int i=0;
	     		int Top=100;
	     		
	     		while (readerAPTList.Read())
	        	{
	     			
	     			if(!readerAPTList.IsDBNull(0)) int_APT_ID = readerAPTList.GetInt32(0);
	       			if(!readerAPTList.IsDBNull(1)) string_APT_ICAO = readerAPTList.GetString(1);
	       			if(!readerAPTList.IsDBNull(2)) string_APT_IATA = readerAPTList.GetString(2);
	       			if(!readerAPTList.IsDBNull(3)) string_APT_LH = readerAPTList.GetString(3);
	       			if(!readerAPTList.IsDBNull(4)) string_APT_FedEx = readerAPTList.GetString(4);
	       			if(!readerAPTList.IsDBNull(5)) string_APT_Charters = readerAPTList.GetString(5);
	       			       			
					Label lbl_APT_ICAO = new Label();
					FontFamily family = new FontFamily("Courier New");
					lbl_APT_ICAO.Font = new Font(family, 11.0f, FontStyle.Bold);
					lbl_APT_ICAO.Tag ="dispose";
	     			lbl_APT_ICAO.Top = Top + 20 * i;
	     			lbl_APT_ICAO.Size = new Size(45, 16);
	     			lbl_APT_ICAO.ForeColor = Color.OrangeRed;
	     			lbl_APT_ICAO.Text = string_APT_ICAO;
	     			lbl_APT_ICAO.Left = 28;
	     			APT_List.Controls.Add(lbl_APT_ICAO);
	     			
	     			Label lbl_APT_IATA = new Label();
					lbl_APT_IATA.Font = new Font(family, 11.0f, FontStyle.Bold);
					lbl_APT_IATA.Tag ="dispose";
	     			lbl_APT_IATA.Top = Top + 20 * i;
	     			lbl_APT_IATA.Size = new Size(65, 16);
	     			lbl_APT_IATA.ForeColor = Color.CornflowerBlue;
	     			lbl_APT_IATA.Text = " - " + string_APT_IATA;
	     			lbl_APT_IATA.Left = 65;
	     			APT_List.Controls.Add(lbl_APT_IATA);
	     			    			
	     			CheckBox ChckBx_APT_Station_LH = new CheckBox();
	     			ChckBx_APT_Station_LH.Enabled = false;
	     			ChckBx_APT_Station_LH.Tag ="dispose";
	     			ChckBx_APT_Station_LH.Top = Top + 20 * i;
	     			ChckBx_APT_Station_LH.Size = new Size(20, 16);
	     			ChckBx_APT_Station_LH.ForeColor = Color.DimGray;
	     			if(string_APT_LH=="Yes") ChckBx_APT_Station_LH.Checked = true;
	     			else ChckBx_APT_Station_LH.Checked = false;
	     			ChckBx_APT_Station_LH.Left = 150;
	     			APT_List.Controls.Add(ChckBx_APT_Station_LH);
	     			
	     			CheckBox ChckBx_APT_Station_FedEx = new CheckBox();
	     			ChckBx_APT_Station_FedEx.Enabled = false;
	     			ChckBx_APT_Station_FedEx.Tag ="dispose";
	     			ChckBx_APT_Station_FedEx.Top = Top + 20 * i;
	     			ChckBx_APT_Station_FedEx.Size = new Size(20, 16);
	     			ChckBx_APT_Station_FedEx.ForeColor = Color.DimGray;
	     			if(string_APT_FedEx=="Yes") ChckBx_APT_Station_FedEx.Checked = true;
	     			else ChckBx_APT_Station_FedEx.Checked = false;
	     			ChckBx_APT_Station_FedEx.Left = 190;
	     			APT_List.Controls.Add(ChckBx_APT_Station_FedEx);
	     			
	     			CheckBox ChckBx_APT_Station_Charters = new CheckBox();
	     			ChckBx_APT_Station_Charters.Enabled = false;
	     			ChckBx_APT_Station_Charters.Tag ="dispose";
	     			ChckBx_APT_Station_Charters.Top = Top + 20 * i;
	     			ChckBx_APT_Station_Charters.Size = new Size(20, 16);
	     			ChckBx_APT_Station_Charters.ForeColor = Color.DimGray;
	     			if(string_APT_Charters=="Yes") ChckBx_APT_Station_Charters.Checked = true;
	     			else ChckBx_APT_Station_Charters.Checked = false;
	     			ChckBx_APT_Station_Charters.Left = 230;
	     			APT_List.Controls.Add(ChckBx_APT_Station_Charters);
	     			     		   			
	     			del_APT_Buttons[int_APT_ID] = new Button();
	     			del_APT_Buttons[int_APT_ID].Tag = "dispose";
	     			del_APT_Buttons[int_APT_ID].Size = new Size(35 , 20);
	        		del_APT_Buttons[int_APT_ID].Location = new Point(310, Top-3+20 * i);
	        		int newSize =7;
	        		del_APT_Buttons[int_APT_ID].Text = "Del";
	        		del_APT_Buttons[int_APT_ID].BackColor = Color.Red;
	        		del_APT_Buttons[int_APT_ID].Font = new Font(del_APT_Buttons[int_APT_ID].Font.FontFamily, newSize);
	        		int i_APT_Del = int_APT_ID;
	        		del_APT_Buttons[int_APT_ID].Click += (sender1, ex) => this.Delete_APT(i_APT_Del);       		
	     			APT_List.Controls.Add(del_APT_Buttons[int_APT_ID]);
	     			
	     			edit_APT_Buttons[int_APT_ID] = new Button();
	     			edit_APT_Buttons[int_APT_ID].Tag = "dispose";
	     			edit_APT_Buttons[int_APT_ID].Size = new Size(35, 20);
	        		edit_APT_Buttons[int_APT_ID].Location = new Point(270, Top-3+20 * i);
	        		edit_APT_Buttons[int_APT_ID].Text = "Edit";
	        		edit_APT_Buttons[int_APT_ID].BackColor = Color.LightBlue;
	        		edit_APT_Buttons[int_APT_ID].Font = new Font(edit_APT_Buttons[int_APT_ID].Font.FontFamily, newSize);
	        		int i_APT_Edit = int_APT_ID;
	        		edit_APT_Buttons[int_APT_ID].Click += (sender1, ex) => this.Edit_APT(i_APT_Edit);
	     			APT_List.Controls.Add(edit_APT_Buttons[int_APT_ID]);
	     					
	     			i++;
	     		}
	     		conn.Close();
			}
		}
		void Delete_APT(int i)
        {
			DialogResult dialogResult = MessageBox.Show("Are you sure that you want to delete ?", "Delete Airport ", MessageBoxButtons.YesNo);
			if(dialogResult == DialogResult.Yes)
			{
            // Connexion à la DB message.mdb
			//try{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
			string insertlog = "DELETE From Stations_ICAO_IATA WHERE ID="+i+"";

			OleDbCommand commandeinsert = new OleDbCommand(insertlog, conn);
			// Execution
							
			commandeinsert.ExecuteNonQuery();
				
			conn.Close();

				//}
//				catch(Exception Ex)
//				{
//					MessageBox.Show("Could not update the database. Your record has not been saved. If error persist, contact the administrator.", "Access database issue", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//				}
			db_read();
				 //do something
			}
			else if (dialogResult == DialogResult.No)
			{
    				//do something else
			}
        }
		void Edit_APT(int i)
        {
			string stringICAO ="";
     		string stringIATA ="";
     		string stringLongHaul="";
			string stringFedEx="";
			string stringCharters ="";
			
            // Connexion à la DB message.mdb
			//try{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
			var query2 = "Select * From Stations_ICAO_IATA WHERE ID="+i+"";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader reader2 = command4.ExecuteReader();
     		
			while (reader2.Read())
        	{
       				if(!reader2.IsDBNull(1)) stringICAO = reader2.GetString(1);
       				if(!reader2.IsDBNull(2)) stringIATA = reader2.GetString(2);
       				if(!reader2.IsDBNull(3)) stringLongHaul = reader2.GetString(3);
       				if(!reader2.IsDBNull(4)) stringFedEx = reader2.GetString(4);
       				if(!reader2.IsDBNull(5)) stringCharters = reader2.GetString(5);
			}
			conn.Close();
			//}
//			catch(Exception Ex)
//			{
//				MessageBox.Show("Could not read the database. If error persist, contact the administrator.", "Access database issue", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//			}	
			TxtBox_APT_ICAO.Text= stringICAO;
			TxtBox_APT_IATA.Text = stringIATA;
			
			if(stringLongHaul=="Yes")ChckBx_APT_LH.Checked=true;
			else ChckBx_APT_LH.Checked = false;
			if(stringFedEx=="Yes")ChckBx_APT_FedEx.Checked=true;
			else ChckBx_APT_FedEx.Checked = false;
			if(stringCharters=="Yes")ChckBx_APT_Charters.Checked=true;
			else ChckBx_APT_Charters.Checked = false;
			
			Btn_addAPT.Text="Edit";
			Btn_addAPT.Tag=i.ToString();
        }
		void Btn_addAPTClick(object sender, EventArgs e)
		{
			string stringICAO = TxtBox_APT_ICAO.Text;
			string stringIATA = TxtBox_APT_IATA.Text;
			string stringLongHaul ="";
			if (ChckBx_APT_LH.Checked)stringLongHaul = "Yes";
			else stringLongHaul="No";
			string stringFedEx ="";
			if (ChckBx_APT_FedEx.Checked)stringFedEx = "Yes";
			else stringFedEx="No";
			string stringCharters ="";
			if (ChckBx_APT_Charters.Checked)stringCharters = "Yes";
			else stringCharters="No";
			
			// Connexion à la DB message.mdb
			try{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
			if(Btn_addAPT.Text=="Edit")
			{
				Btn_addAPT.Text="Add Airport !";
				string editID=Btn_addAPT.Tag.ToString();
				int intID=int.Parse(editID);
				var updatelog = "UPDATE Stations_ICAO_IATA SET ICAO='"+stringICAO+"',IATA='"+stringIATA+"'," +
					"LH='"+stringLongHaul+"',FedEx='"+stringFedEx+"',Charters='"+stringCharters+"' WHERE ID="+intID+"";
				
				OleDbCommand commandeinsert = new OleDbCommand(updatelog, conn);
				// Execution
				commandeinsert.ExecuteNonQuery();
			}
			else
			{
				var insertlog = "INSERT INTO Stations_ICAO_IATA ([ICAO], [IATA], [LH], [FedEx], [Charters]) VALUES" +
"					('"+stringICAO+"','"+stringIATA+"','"+stringLongHaul+"','"+stringFedEx+"','"+stringCharters+"')";

				//var insertlog = "INSERT INTO morningTable(db_dlaFx) Values ('"+dlaFx+"')";
				OleDbCommand commandeinsert = new OleDbCommand(insertlog, conn);
				// Execution
				
				
				commandeinsert.ExecuteNonQuery();
			}
				
				conn.Close();
				}
				catch(Exception Ex)
				{
					MessageBox.Show("Could not update the database. Your record has not been saved. If error persist, contact the administrator.", "Access database issue", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			
			
				db_read();
		}	

		void Btn_CopyAPTListClick(object sender, EventArgs e)
		{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
       		var queryAPTList = "Select * From Stations_ICAO_IATA ORDER BY ICAO";
     		OleDbCommand commandAPTList = new OleDbCommand(queryAPTList, conn);
     		OleDbDataReader readerAPTList = commandAPTList.ExecuteReader();
     		
     		string stringAPTList="";
     		
     		while (readerAPTList.Read())
        	{
       			if(!readerAPTList.IsDBNull(1)) stringAPTList += readerAPTList.GetString(1)+ " ";
     		}
		
     		Clipboard.SetText(stringAPTList);
		}
		void Btn_CopyAPTList2Click(object sender, EventArgs e)
		{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
       		var queryAPTList = "Select * From Stations_ICAO_IATA ORDER BY ICAO";
     		OleDbCommand commandAPTList = new OleDbCommand(queryAPTList, conn);
     		OleDbDataReader readerAPTList = commandAPTList.ExecuteReader();
     		
     		string stringAPTList="";
     		
     		while (readerAPTList.Read())
        	{
       			if(!readerAPTList.IsDBNull(1)) stringAPTList += readerAPTList.GetString(1)+ " ";
     		}
		
     		Clipboard.SetText(stringAPTList);
		}

		void Btn_refreshAppClick(object sender, EventArgs e)
		{
			db_read();
		}
		void Btn_editListClick(object sender, EventArgs e)
		{
			if(btn_editList.Text=="Edit List")
			{
				btn_editList.Text="Close List";
				db_read();
			}
			else if(btn_editList.Text=="Close List")
			{
				btn_editList.Text="Edit List";
				db_read();
			}
		}
	}
}
