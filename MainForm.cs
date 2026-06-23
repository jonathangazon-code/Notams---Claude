/*
 * Created by SharpDevelop.
 * User: jgazon
 * Date: 24-02-21
 * Time: 17:16
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Linq;

namespace ICAO_CSV
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

			StartApp();
			LoadStationsCache();
			//Report();
			tab_RWYs();
			Airport_List();
			AIP_SUP_Checklist();
			
			// 👉 Appelé automatiquement quand on ferme l’application
	        this.FormClosing += MainForm_FormClosing;

			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		
		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		    {
		        EndApp();
		    }


		public void StartApp()		
		{
		    // Chemin source (drive V:)
		    string sourcePath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\NOTAMS APP\ICAO_storedNotams.mdb";
		
		    // Dossier local = dossier du .exe
		    string appFolder = Application.StartupPath;
		
		    // Destination
		    string destPath = Path.Combine(appFolder, "ICAO_storedNotams.mdb");
		
		    try
		    {
		        // Vérification que la source existe
		        if (!File.Exists(sourcePath))
		        {
		            MessageBox.Show(
		                "Fichier source introuvable :\n" + sourcePath,
		                "Erreur",
		                MessageBoxButtons.OK,
		                MessageBoxIcon.Error);
		            return;
		        }
		
		        // Si le fichier local existe déjà → on compare les dates
		        if (File.Exists(destPath))
		        {
		            DateTime sourceDate = File.GetLastWriteTime(sourcePath);
		            DateTime destDate = File.GetLastWriteTime(destPath);
		
		            // Si la version locale est à jour → ne rien faire
		            if (destDate >= sourceDate)
		            {
		                MessageBox.Show(
		                    "Local DB already up to date.\n\nNo download needed.",
		                    "Information",
		                    MessageBoxButtons.OK,
		                    MessageBoxIcon.Information);
		                return;
		            }
		        }
		
		        // Sinon → copier (overwrite)
		        File.Copy(sourcePath, destPath, true);
		
		        MessageBox.Show(
		            "DB Downloaded\n\nSaved to:\n" + destPath,
		            "OK",
		            MessageBoxButtons.OK,
		            MessageBoxIcon.Information);
		    }
		    catch (Exception ex)
		    {
		        MessageBox.Show(
		            "Erreur lors de la copie :\n" + ex.Message,
		            "Erreur",
		            MessageBoxButtons.OK,
		            MessageBoxIcon.Error);
		    }
		}	
		
		
		public void EndApp()
		{
		    // Chemin local = dossier de l'executable
		    string appFolder = Application.StartupPath;
		    string localPath = Path.Combine(appFolder, "ICAO_storedNotams.mdb");
		
		    // Chemin sur le drive V:
		    string vDrivePath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\NOTAMS APP\ICAO_storedNotams.mdb";
		
		    try
		    {
		        // Vérifie que le fichier local existe
		        if (!File.Exists(localPath))
		        {
		            MessageBox.Show(
		                "Local database not found:\n" + localPath + 
		                "\n\nUpload to V: cancelled.",
		                "Info",
		                MessageBoxButtons.OK,
		                MessageBoxIcon.Information);
		            return;
		        }
		
		        // Si le fichier V: n'existe pas → copie directe
		        if (!File.Exists(vDrivePath))
		        {
		            File.Copy(localPath, vDrivePath, true);
		            MessageBox.Show(
		                "DB uploaded to V: (file created).\n\n" + vDrivePath,
		                "Success",
		                MessageBoxButtons.OK,
		                MessageBoxIcon.Information);
		            return;
		        }
		
		        // Comparer les dates
		        DateTime localDate = File.GetLastWriteTime(localPath);
		        DateTime vDate = File.GetLastWriteTime(vDrivePath);
		
		        // Copier si local plus récent
		        if (localDate > vDate)
		        {
		            File.Copy(localPath, vDrivePath, true);
		            MessageBox.Show(
		                "DB on V: updated (local version was newer).",
		                "Updated",
		                MessageBoxButtons.OK,
		                MessageBoxIcon.Information);
		        }
		        else
		        {
		            MessageBox.Show(
		                "No upload needed.\nV: is up to date.",
		                "Info",
		                MessageBoxButtons.OK,
		                MessageBoxIcon.Information);
		        }
		    }
		    catch (Exception ex)
		    {
		        MessageBox.Show(
		            "Error uploading DB to V:\n" + ex.Message,
		            "Error",
		            MessageBoxButtons.OK,
		            MessageBoxIcon.Error);
		    }
		}


		
		void update_SUP_Avio(int i)
		{
			string num_ID="";
			num_ID=i.ToString();
			
			RchTxtBox_Test.Text="le num ID est : "+num_ID;
			
			string status="";
			string newStatus="";
			
			System.Data.OleDb.OleDbConnection conn2 = new System.Data.OleDb.OleDbConnection();
			conn2.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
			conn2.Open();
			
			var query2 = "SELECT * FROM filteredNotams_table WHERE ID=?";
     		OleDbCommand command4 = new OleDbCommand(query2, conn2);
     		command4.Parameters.AddWithValue("?", i);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
			
			while (dBreader.Read())
            	{
            	
				if(!dBreader.IsDBNull(15)) status = dBreader.GetString(15);
				}
			
			conn2.Close();
			
			if(status=="Yes") newStatus="";
			else newStatus="Yes";
			
			RchTxtBox_Test.Text+=" "+newStatus;
			
			conn2.Open();
		
			
			var update = "UPDATE filteredNotams_table SET Loaded_Aviobook=? WHERE ID=?";
     		OleDbCommand commandUpdate = new OleDbCommand(update, conn2);
     		commandUpdate.Parameters.AddWithValue("?", newStatus);
     		commandUpdate.Parameters.AddWithValue("?", i);
     		commandUpdate.ExecuteNonQuery();
			
			conn2.Close();
			
			AIP_SUP_Checklist();
		}
		
		public void AIP_SUP_Checklist()
		{
			
			List<Label> itemsToRemove = new List<Label>();
			foreach (Label label in AIP_Sup.Controls.OfType<Label>())
			{
    			if (label.Tag != null && label.Tag.ToString() == "dispose")
    			{
      		  		itemsToRemove.Add(label);
    			}
			}
			foreach (Label label in itemsToRemove)
			{
    			AIP_Sup.Controls.Remove(label);
    			label.Dispose();
			}
			
			List<CheckBox> chckboxToRemove = new List<CheckBox>();
			foreach (CheckBox chckbox in AIP_Sup.Controls.OfType<CheckBox>())
			{
    			if (chckbox.Tag != null && chckbox.Tag.ToString() == "dispose")
    			{
      		  		chckboxToRemove.Add(chckbox);
    			}
			}
			foreach (CheckBox chckbox in chckboxToRemove)
			{
    			AIP_Sup.Controls.Remove(chckbox);
    			chckbox.Dispose();
			}
			
			AIP_Sup.AutoScroll=true;
			AIP_Sup.VerticalScroll.Value = 0;
			int baseTop=10;
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM filteredNotams_table ORDER BY location";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
     		string fromDate="";
			string tillDate="";
			//string all="";
			string ICAO="";
			string location="";
			string key="";
			string Status="";
			string Impact="";
			string Remark="";
			int int_APT_ID=0;
			string AIP_SUP_List="";
			
			int i=0;
			
			CheckBox[] ChckBx_Loaded_Aviobook = new CheckBox[32000];
			
			if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		string loaded_aviobook="";
            		if(!dBreader.IsDBNull(12))
            		{
            			
            			Impact="";
            			Status="";
            			Remark="";
            			fromDate="";
            			tillDate="";
            			
            			Status=dBreader.GetString(12);
            			if(Status=="K")
            			{
            				if(!dBreader.IsDBNull(13)) Impact = dBreader.GetString(13);
            				if(Impact=="AS")
            				{
	            				if(!dBreader.IsDBNull(5)) fromDate = dBreader.GetString(5);
	            				if(!dBreader.IsDBNull(6)) tillDate = dBreader.GetString(6);
	            				string fromDateYear= fromDate.Substring(0,4);
	            				string fromDateMonth=fromDate.Substring(5,2);
	            				string fromDateDay=fromDate.Substring(8,2);
	     						string tillDateYear= tillDate.Substring(0,4);
	     						string tillDateMonth=tillDate.Substring(5,2);
	     						string tillDateDay=tillDate.Substring(8,2);
	     						
	     						string fromDateCheck=fromDateYear+fromDateMonth+fromDateDay;
	     						string tillDateCheck=tillDateYear+tillDateMonth+tillDateDay;
	     						
	     						
	     						if(!dBreader.IsDBNull(0)) int_APT_ID = dBreader.GetInt32(0);
	     						if(!dBreader.IsDBNull(8)) ICAO = dBreader.GetString(8);
			            		//if(!dBreader.IsDBNull(7)) all = dBreader.GetString(7);
			            		if(!dBreader.IsDBNull(10)) key = dBreader.GetString(10);
			            		
			            		if(!dBreader.IsDBNull(14)) Remark = dBreader.GetString(14);
			            		if(!dBreader.IsDBNull(15)) loaded_aviobook = dBreader.GetString(15);
			            				
			            		location = GetIATA(ICAO);
			            				
			            		fromDate = dateTransformation(fromDate);
			            		tillDate = dateTransformation(tillDate);
			            		
			            		//AIP_SUP_List+=ICAO+" "+key+" "+fromDate+"-"+tillDate+" "+Remark+"\n";
			            		
			            		Label lbl_APT_IATA = new Label();
								FontFamily family = new FontFamily("Courier New");
								lbl_APT_IATA.Font = new Font(family, 11.0f, FontStyle.Bold);
								lbl_APT_IATA.Tag ="dispose";
     							lbl_APT_IATA.Top = baseTop + 20 * i;
     							lbl_APT_IATA.Size = new Size(45, 16);
     							lbl_APT_IATA.ForeColor = Color.BlueViolet;
     							lbl_APT_IATA.Text = location;
     							lbl_APT_IATA.Left = 28;
     							AIP_Sup.Controls.Add(lbl_APT_IATA);
     							
     							Label lbl_key = new Label();
								//FontFamily family = new FontFamily("Courier New");
								lbl_key.Font = new Font(family, 11.0f, FontStyle.Bold);
								lbl_key.Tag ="dispose";
     							lbl_key.Top = baseTop + 20 * i;
     							lbl_key.Size = new Size(130, 16);
     							lbl_key.ForeColor = Color.Black;
     							lbl_key.Text = key;
     							lbl_key.Left = 80;
     							AIP_Sup.Controls.Add(lbl_key);
     							
     							string dates="";
     							dates=fromDate+"-"+tillDate;
     							
     							Label lbl_dates = new Label();
								//FontFamily family = new FontFamily("Courier New");
								lbl_dates.Font = new Font(family, 11.0f, FontStyle.Bold);
								lbl_dates.Tag ="dispose";
     							lbl_dates.Top = baseTop + 20 * i;
     							lbl_dates.Size = new Size(200, 16);
     							lbl_dates.ForeColor = Color.Black;
     							lbl_dates.Text = dates;
     							lbl_dates.Left = 230;
     							AIP_Sup.Controls.Add(lbl_dates);
     							
     							Label lbl_remark = new Label();
								//FontFamily family = new FontFamily("Courier New");
								lbl_remark.Font = new Font(family, 11.0f, FontStyle.Bold);
								lbl_remark.Tag ="dispose";
     							lbl_remark.Top = baseTop + 20 * i;
     							lbl_remark.Size = new Size(300, 16);
     							lbl_remark.ForeColor = Color.Black;
     							lbl_remark.Text = Remark;
     							lbl_remark.Left = 450;
     							AIP_Sup.Controls.Add(lbl_remark);
								
								ChckBx_Loaded_Aviobook[int_APT_ID] = new CheckBox();
     							//ChckBx_Loaded_Aviobook[int_APT_ID].Enabled = false;
     							ChckBx_Loaded_Aviobook[int_APT_ID].Tag ="dispose";
     							ChckBx_Loaded_Aviobook[int_APT_ID].Top = baseTop + 20 * i;
     							ChckBx_Loaded_Aviobook[int_APT_ID].Size = new Size(20, 16);
     							
     							if(loaded_aviobook=="Yes") 
     							{
     								ChckBx_Loaded_Aviobook[int_APT_ID].Checked = true;
     								ChckBx_Loaded_Aviobook[int_APT_ID].BackColor = Color.DimGray;
     							}
     							else
     							{
     								ChckBx_Loaded_Aviobook[int_APT_ID].Checked = false;
     								ChckBx_Loaded_Aviobook[int_APT_ID].BackColor = Color.Red;
     							}
     							ChckBx_Loaded_Aviobook[int_APT_ID].Left = 430;
     							AIP_Sup.Controls.Add(ChckBx_Loaded_Aviobook[int_APT_ID]);
     							
     							int i_APT_change = int_APT_ID;
				        		ChckBx_Loaded_Aviobook[int_APT_ID].Click += (sender1, ex) => this.update_SUP_Avio(i_APT_change);
     							
     							i++;
            				}
            			}
            		}
            	}
			}
			conn.Close();		
		}
		
		public void Airport_List()
		{
			
			//APT_List.VerticalScroll.Value = 0;
			
			List<Label> itemsToRemove = new List<Label>();
			foreach (Label label in APT_List.Controls.OfType<Label>())
			{
    			if (label.Tag != null && label.Tag.ToString() == "dispose")
    			{
      		  		itemsToRemove.Add(label);
    			}
			}
			foreach (Label label in itemsToRemove)
			{
    			APT_List.Controls.Remove(label);
    			label.Dispose();
			}
			
			List<TextBox> txtboxToRemove = new List<TextBox>();
			foreach (TextBox txtbox in APT_List.Controls.OfType<TextBox>())
			{
    			if (txtbox.Tag != null && txtbox.Tag.ToString() == "dispose")
    			{
      		  		txtboxToRemove.Add(txtbox);
    			}
			}
			foreach (TextBox txtbox in txtboxToRemove)
			{
    			APT_List.Controls.Remove(txtbox);
    			txtbox.Dispose();
			}
			
			List<RichTextBox> rchtxtboxToRemove = new List<RichTextBox>();
			foreach (RichTextBox rchtxtbox in APT_List.Controls.OfType<RichTextBox>())
			{
    			if (rchtxtbox.Tag != null && rchtxtbox.Tag.ToString() == "dispose")
    			{
      		  		rchtxtboxToRemove.Add(rchtxtbox);
    			}
			}
			foreach (RichTextBox rchtxtbox in rchtxtboxToRemove)
			{
    			APT_List.Controls.Remove(rchtxtbox);
    			rchtxtbox.Dispose();
			}
			
			List<CheckBox> chckboxToRemove = new List<CheckBox>();
			foreach (CheckBox chckbox in APT_List.Controls.OfType<CheckBox>())
			{
    			if (chckbox.Tag != null && chckbox.Tag.ToString() == "dispose")
    			{
      		  		chckboxToRemove.Add(chckbox);
    			}
			}
			foreach (CheckBox chckbox in chckboxToRemove)
			{
    			APT_List.Controls.Remove(chckbox);
    			chckbox.Dispose();
			}
			
			List<Button> buttonsToRemove = new List<Button>();
			foreach (Button button in APT_List.Controls.OfType<Button>())
			{
    			if (button.Tag != null && button.Tag.ToString() == "dispose")
    			{
      		  		buttonsToRemove.Add(button);
    			}
			}
			foreach (Button button in buttonsToRemove)
			{
    			APT_List.Controls.Remove(button);
    			button.Dispose();
			}
			
			APT_List.AutoScroll=true;
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
				
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
     		
     		Button[] del_APT_Buttons = new Button[3200];
     		Button[] edit_APT_Buttons = new Button[3200];
     		int i=0;
     		Top=70;
     		
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
			string insertlog = "DELETE From Stations_ICAO_IATA WHERE ID=?";

			OleDbCommand commandeinsert = new OleDbCommand(insertlog, conn);
			commandeinsert.Parameters.AddWithValue("?", i);
			// Execution

			commandeinsert.ExecuteNonQuery();

			conn.Close();
			LoadStationsCache();

				//}
//				catch(Exception Ex)
//				{
//					MessageBox.Show("Could not update the database. Your record has not been saved. If error persist, contact the administrator.", "Access database issue", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//				}
			//db_read();
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
			var query2 = "Select * From Stations_ICAO_IATA WHERE ID=?";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		command4.Parameters.AddWithValue("?", i);
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
		
		public static string dateTransformation(string notamDate)
		{
			string reportDate ="";
			string notamHour ="";
			string notamDay ="";
			string notamMonth ="";
			string notamYear="";
			
			notamDate=notamDate.Substring(0,16);
     		notamHour=notamDate.Substring(11,5);
     		notamDay=notamDate.Substring(8,2);
     		notamMonth=notamDate.Substring(5,2);
     		if(notamMonth=="01")notamMonth="JAN";
     		if(notamMonth=="02")notamMonth="FEB";
     		if(notamMonth=="03")notamMonth="MAR";
     		if(notamMonth=="04")notamMonth="APR";
     		if(notamMonth=="05")notamMonth="MAY";
     		if(notamMonth=="06")notamMonth="JUN";
     		if(notamMonth=="07")notamMonth="JUL";
     		if(notamMonth=="08")notamMonth="AUG";
     		if(notamMonth=="09")notamMonth="SEP";
     		if(notamMonth=="10")notamMonth="OCT";
     		if(notamMonth=="11")notamMonth="NOV";
     		if(notamMonth=="12")notamMonth="DEC";
     		notamYear=notamDate.Substring(0,4);
     		//reportDate=notamDay+notamMonth+notamYear+" "+notamHour;
     		reportDate=notamDay+notamMonth+notamYear;
     		
     		return reportDate;
		}
		public void Sup_Report()
		{
			string fromDate="";
			string tillDate="";
			string all="";
			string ICAO="";
			string location="";
			string key="";
			string Status="";
			string Impact="";
			string Remark="";
			
			string twenty4H_fromDate="";
			string twenty4H_tillDate="";
			string twenty4H_all="";
			string twenty4H_ICAO="";
			string twenty4H_location="";
			string twenty4H_key="";
			string twenty4H_Status="";
			string twenty4H_Impact="";
			string twenty4H_Remark="";
			
			string AIP_Sup_list="";
			string twenty4H_AIP_Sup_list="";	
			
			string report ="";
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM filteredNotams_table ORDER BY location";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		//*******************************************************************
     		string output="";
     		string test="";
     		string result="";
     		string impactTest="";
     		
     		
     		if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		string loaded_aviobook="";
//            		test=dBreader.GetString(8);
//            		
//     
//            		
//            		if(test=="ZSSS")
//            		{
//            			if(!dBreader.IsDBNull(13))
//            			{
//            				//output="test";
//            				impactTest="0";
//            				impactTest=dBreader.GetString(13);
//            				//output=impactTest;
//            				if (impactTest!="C") output+="T";
//            				if(impactTest=="C") output+=impactTest;
//            				
//            			}
//            		}
            		
            		
            		Impact="";
            		Status="";
            		Remark="";
            		fromDate="";
            		tillDate="";
            //**********************************************************************		
            		if(!dBreader.IsDBNull(12))
            		{
            			Status=dBreader.GetString(12);
            			if(Status=="K")
            			{
            				if(!dBreader.IsDBNull(5)) fromDate = dBreader.GetString(5);
            				if(!dBreader.IsDBNull(6)) tillDate = dBreader.GetString(6);
            				string fromDateYear= fromDate.Substring(0,4);
            				string fromDateMonth=fromDate.Substring(5,2);
            				string fromDateDay=fromDate.Substring(8,2);
     						string tillDateYear= tillDate.Substring(0,4);
     						string tillDateMonth=tillDate.Substring(5,2);
     						string tillDateDay=tillDate.Substring(8,2);
     						
     						string fromDateCheck=fromDateYear+fromDateMonth+fromDateDay;
     						string tillDateCheck=tillDateYear+tillDateMonth+tillDateDay;
     						
            				int fromDateInt= Int32.Parse(fromDateCheck);
            				int tillDateInt= Int32.Parse(tillDateCheck);
            				
            				DateTime today = DateTime.Now;
        					DateTime tomorrow = today.AddDays(1);
        					DateTime sevenDays = today.AddDays(7);
        					DateTime thirtyOneDays = today.AddDays(31);
            				
        					
        					
            				var todayString= DateTime.Now.ToString("yyyyMMdd");
            				var tomorrowString= tomorrow.ToString("yyyyMMdd");
            				var sevenDaysString= sevenDays.ToString("yyyyMMdd");
            				var thirtyOneDaysString= thirtyOneDays.ToString("yyyyMMdd");
            				
            				int tomorrowInt=Int32.Parse(tomorrowString);
            				            				
            				int todayInt=0;
            				int endWindowInt=30000101;
            				
            				
            				if(!dBreader.IsDBNull(15)) loaded_aviobook = dBreader.GetString(15);
            				
            				if(radBtn_Sup_24Hrs.Checked==true)
            				{
            					todayInt= Int32.Parse(todayString);
            					endWindowInt= Int32.Parse(tomorrowString);
            					//RchTxtCSV.Text=tomorrowString;
            				}
            				
            				if(radBtn_Sup_7days.Checked==true)
            				{
            					todayInt= Int32.Parse(todayString);
            					endWindowInt= Int32.Parse(sevenDaysString);
            					//RchTxtCSV.Text=sevenDaysString;
            				}
            				
            				if(radBtn_Sup_31days.Checked==true)
            				{
            					todayInt= Int32.Parse(todayString);
            					endWindowInt= Int32.Parse(thirtyOneDaysString);
            					//RchTxtCSV.Text=thirtyOneDaysString;
            				}       					
            				         				
            				if(tillDateInt>todayInt && fromDateInt<endWindowInt)
            				{
            					string LHtype="LH";
            					string SHtype="FedEx";
            					string Chartertype="Charters";
            					
            					string fromTest= fromDateInt.ToString();
            					string tomorrowTest=tomorrowInt.ToString();
            					//RchTxtCSV.Text=fromTest+"<br />"+tomorrowTest;
            					
            					string checkbox_status="";
		            				if(loaded_aviobook=="Yes")
		            				{
		            					checkbox_status="<th style=\"width:30px\"><input type=\"checkbox\" checked></th>";
		            				}
		            				else checkbox_status="<th bgcolor=\"Red\" style=\"width:30px\"><input type=\"checkbox\"></th>";
		            				            					
            					//Next 24 hrs
            					//-----------
            					if(fromDateInt<=tomorrowInt)
            					{
            						
            						if(!dBreader.IsDBNull(8)) ICAO = dBreader.GetString(8);
		            				if(!dBreader.IsDBNull(7)) all = dBreader.GetString(7);
		            				if(!dBreader.IsDBNull(10)) key = dBreader.GetString(10);
		            				if(!dBreader.IsDBNull(13)) Impact = dBreader.GetString(13);
		            				if(!dBreader.IsDBNull(14)) Remark = dBreader.GetString(14);
		            				
		            				//location = "<span style=\"color: blue;font-weight:bold\">"+ ICAO +" - "+GetIATA(ICAO) + " : </span>";
		            				location = GetIATA(ICAO);
		            				
		            				fromDate = dateTransformation(fromDate);
		            				tillDate = dateTransformation(tillDate);
		            				
		            				
		            					
			            			if(Impact=="AS") twenty4H_AIP_Sup_list+="<tr><th bgcolor=\"Yellow\" style=\"color:SaddleBrown;\">"+location+"</th><th style=\"width:100px;font-family:Courier New;\">"+key+
			            				"</th><th style=\"width:140px; font-family:Courier New;padding-right:10px;\">"+fromDate+"-"+tillDate+
			            				"</th>"+checkbox_status+"<th bgcolor=\"Yellow\" style=\"font-weight:normal;\">"+Remark+"</th></tr>";
			            		
            					}
            					
		            			//Beyond 24 hrs
		            			//------------
		            			else
            					{
		            				if(!dBreader.IsDBNull(8)) ICAO = dBreader.GetString(8);
		            				if(!dBreader.IsDBNull(7)) all = dBreader.GetString(7);
		            				if(!dBreader.IsDBNull(10)) key = dBreader.GetString(10);
		            				if(!dBreader.IsDBNull(13)) Impact = dBreader.GetString(13);
		            				if(!dBreader.IsDBNull(14)) Remark = dBreader.GetString(14);
		            				
		            				//location = "<span style=\"color: blue;font-weight:bold\">"+ ICAO +" - "+GetIATA(ICAO) + " : </span>";
		            				location = GetIATA(ICAO);
		            				
		            				fromDate = dateTransformation(fromDate);
		            				tillDate = dateTransformation(tillDate);
		            				
		            				
		            				
			            			if(Impact=="AS") AIP_Sup_list+="<tr><th style=\"font-weight:bold;color:SaddleBrown;\">"+location+"</th><th style=\"width:100px;font-family:Courier New;\">"+key+
			            				"</th><th style=\"width:140px; font-family:Courier New;padding-right:10px;\">"+fromDate+"-"+tillDate+
			            				"</th>"+checkbox_status+"<th style=\"font-weight:normal;\">"+Remark+"</th></tr>";
			            				
            					}
		            			//RchTxtCSV.Text=twenty4H_APClsdCharters;
            				}
            			}
            		}
            	}
        	}
     		RchTxt_FilterNotams.Text=output;
     		
     		// *********************
     		string window="";
     		if(radBtn_Sup_noFilter.Checked)window="COMPLETE";
     		if(radBtn_Sup_24Hrs.Checked)window="Next 24 Hours";
     		if(radBtn_Sup_7days.Checked)window="Next 7 days";
     		if(radBtn_Sup_31days.Checked)window="Next 31 days";
     		
     		string reportDate= DateTime.Now.ToString("ddMMMMyyyy HHmm")+"CET";
     		
     		report="<html><head><title>AIP SUP Listing</title><body style=\"font-family:Calibri\">";
     		report+="<h1>AIP SUP Listing - "+window+"</h1>";
     		report+="<p>"+reportDate+"</p>";
     		report+="<table border=\"1\" style=\"width:700px; text-align: left; font-family:Calibri; font-size:12px; border : 1px solid black; border-collapse: collapse\">";
     		report+="<tr bgcolor=\"Black\" style=\"color: white; font-size:14px;\"><th style=\"font-weight:bold; border : 1px solid black;\">" +
				"IATA</span></th><th>Notam Ref</th><th>From-Till</th><th>Avio</th><th>AIP SUP Ref</th></tr>"+
     			twenty4H_AIP_Sup_list+AIP_Sup_list;
     			
     		report+="</table></body></html>";
     		
     		conn.Close();
     		
			Web_Sup_report.DocumentText = report;
			
			// Base directory
	        string directoryPath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports";
	
	        // Create reversed date filename (ddMMyy.html)
	        string fileName = DateTime.Now.ToString("yyMMdd") + "-AIP_SUP_report.html";
	
	        // Full file path
	        string fullPath = Path.Combine(directoryPath, fileName);
	
	        // Ensure directory exists
	        Directory.CreateDirectory(directoryPath);
	
	        // Write file (overwrites if it already exists)
	        File.WriteAllText(fullPath, report);
		}
		
		public void Report()
		{
			string fromDate="";
			string tillDate="";
			string all="";
			string ICAO="";
			string location="";
			string key="";
			string Status="";
			string Impact="";
			string Remark="";
			
			string twenty4H_fromDate="";
			string twenty4H_tillDate="";
			string twenty4H_all="";
			string twenty4H_ICAO="";
			string twenty4H_location="";
			string twenty4H_key="";
			string twenty4H_Status="";
			string twenty4H_Impact="";
			string twenty4H_Remark="";
			
			string APClsdLH="";
			string RWYClsdLH="";
			string CatILH="";
			string NilsLH="";
			string NoAltnLH="";
			string FuelLH="";
			string MiscLH="";
			string APClsdSH="";
			string RWYClsdSH="";
			string CatISH="";
			string NilsSH="";
			string NoAltnSH="";
			string FuelSH="";
			string MiscSH="";
			string APClsdCharters="";
			string RWYClsdCharters="";
			string CatICharters="";
			string NilsCharters="";
			string NoAltnCharters="";
			string FuelCharters="";
			string MiscCharters="";
			
			string twenty4H_APClsdLH="";
			string twenty4H_RWYClsdLH="";
			string twenty4H_CatILH="";
			string twenty4H_NilsLH="";
			string twenty4H_NoAltnLH="";
			string twenty4H_FuelLH="";
			string twenty4H_MiscLH="";
			string twenty4H_APClsdSH="";
			string twenty4H_RWYClsdSH="";
			string twenty4H_CatISH="";
			string twenty4H_NilsSH="";
			string twenty4H_NoAltnSH="";
			string twenty4H_FuelSH="";
			string twenty4H_MiscSH="";
			string twenty4H_APClsdCharters="";
			string twenty4H_RWYClsdCharters="";
			string twenty4H_CatICharters="";
			string twenty4H_NilsCharters="";
			string twenty4H_NoAltnCharters="";
			string twenty4H_FuelCharters="";
			string twenty4H_MiscCharters="";
			
			string report ="";
			string reportLH="";
			string reportSH="";
			string reportCharter="";
			
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM filteredNotams_table ORDER BY LOCATION";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		//*******************************************************************
     		string output="";
     		string test="";
     		string result="";
     		string impactTest="";
     		if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		
//            		test=dBreader.GetString(8);
//            		
//     
//            		
//            		if(test=="ZSSS")
//            		{
//            			if(!dBreader.IsDBNull(13))
//            			{
//            				//output="test";
//            				impactTest="0";
//            				impactTest=dBreader.GetString(13);
//            				//output=impactTest;
//            				if (impactTest!="C") output+="T";
//            				if(impactTest=="C") output+=impactTest;
//            				
//            			}
//            		}
//            		
            		
            		Impact="";
            		Status="";
            		Remark="";
            		fromDate="";
            		tillDate="";
            //**********************************************************************		
            		if(!dBreader.IsDBNull(12))
            		{
            			Status=dBreader.GetString(12);
            			if(Status=="K")
            			{
            				if(!dBreader.IsDBNull(5)) fromDate = dBreader.GetString(5);
            				if(!dBreader.IsDBNull(6)) tillDate = dBreader.GetString(6);
            				string fromDateYear= fromDate.Substring(0,4);
            				string fromDateMonth=fromDate.Substring(5,2);
            				string fromDateDay=fromDate.Substring(8,2);
     						string tillDateYear= tillDate.Substring(0,4);
     						string tillDateMonth=tillDate.Substring(5,2);
     						string tillDateDay=tillDate.Substring(8,2);
     						
     						string fromDateCheck=fromDateYear+fromDateMonth+fromDateDay;
     						string tillDateCheck=tillDateYear+tillDateMonth+tillDateDay;
     						
            				int fromDateInt= Int32.Parse(fromDateCheck);
            				int tillDateInt= Int32.Parse(tillDateCheck);
            				
            				DateTime today = DateTime.Now;
        					DateTime tomorrow = today.AddDays(1);
        					DateTime sevenDays = today.AddDays(7);
        					DateTime thirtyOneDays = today.AddDays(31);
            				
        					
        					
            				var todayString= DateTime.Now.ToString("yyyyMMdd");
            				var tomorrowString= tomorrow.ToString("yyyyMMdd");
            				var sevenDaysString= sevenDays.ToString("yyyyMMdd");
            				var thirtyOneDaysString= thirtyOneDays.ToString("yyyyMMdd");
            				
            				int tomorrowInt=Int32.Parse(tomorrowString);
            				            				
            				int todayInt=0;
            				int endWindowInt=30000101;
            				
            				if(radBtn_24Hrs.Checked==true)
            				{
            					todayInt= Int32.Parse(todayString);
            					endWindowInt= Int32.Parse(tomorrowString);
            					//RchTxtCSV.Text=tomorrowString;
            				}
            				
            				if(radBtn_7days.Checked==true)
            				{
            					todayInt= Int32.Parse(todayString);
            					endWindowInt= Int32.Parse(sevenDaysString);
            					//RchTxtCSV.Text=sevenDaysString;
            				}
            				
            				if(radBtn_31days.Checked==true)
            				{
            					todayInt= Int32.Parse(todayString);
            					endWindowInt= Int32.Parse(thirtyOneDaysString);
            					//RchTxtCSV.Text=thirtyOneDaysString;
            				}       					
            				         				
            				if(tillDateInt>todayInt && fromDateInt<endWindowInt)
            				{
            					string LHtype="LH";
            					string SHtype="FedEx";
            					string Chartertype="Charters";
            					
            					string fromTest= fromDateInt.ToString();
            					string tomorrowTest=tomorrowInt.ToString();
            					//RchTxtCSV.Text=fromTest+"<br />"+tomorrowTest;
            					
            					
            					//Next 24 hrs
            					//-----------
            					if(fromDateInt<=tomorrowInt)
            					{
            						
            						if(!dBreader.IsDBNull(8)) ICAO = dBreader.GetString(8);
		            				if(!dBreader.IsDBNull(7)) all = dBreader.GetString(7);
		            				if(!dBreader.IsDBNull(10)) key = dBreader.GetString(10);
		            				if(!dBreader.IsDBNull(13)) Impact = dBreader.GetString(13);
		            				if(!dBreader.IsDBNull(14)) Remark = dBreader.GetString(14);
		            				
		            				//location = "<span style=\"color: blue;font-weight:bold\">"+ ICAO +" - "+GetIATA(ICAO) + " : </span>";
		            				location = GetIATA(ICAO);
		            				
		            				fromDate = dateTransformation(fromDate);
		            				tillDate = dateTransformation(tillDate);
		            				
		            				
		            				if(IsOpsType(LHtype, ICAO)=="Yes")
		            				{
			            				if(Impact=="A") twenty4H_APClsdLH+="<tr><th bgcolor=\"Yellow\" style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="R") twenty4H_RWYClsdLH+="<tr><th bgcolor=\"Yellow\" style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="C") twenty4H_CatILH+="<tr><th bgcolor=\"Yellow\" style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="N") twenty4H_NilsLH+="<tr><th bgcolor=\"Yellow\" style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="D") twenty4H_NoAltnLH+="<tr><th bgcolor=\"Yellow\" style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="F") twenty4H_FuelLH+="<tr><th bgcolor=\"Yellow\" style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="M") twenty4H_MiscLH+="<tr><th bgcolor=\"Yellow\" style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
		            				}
		            				
		            				if(IsOpsType(SHtype, ICAO)=="Yes")
		            				{
			            				if(Impact=="A") twenty4H_APClsdSH+="<tr><th bgcolor=\"Yellow\" style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="R") twenty4H_RWYClsdSH+="<tr><th bgcolor=\"Yellow\" style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="C") twenty4H_CatISH+="<tr><th bgcolor=\"Yellow\" style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="N") twenty4H_NilsSH+="<tr><th bgcolor=\"Yellow\" style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="D") twenty4H_NoAltnSH+="<tr><th bgcolor=\"Yellow\" style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="F") twenty4H_FuelSH+="<tr><th bgcolor=\"Yellow\" style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="M") twenty4H_MiscSH+="<tr><th bgcolor=\"Yellow\" style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
		            				}
		            				
		            				if(IsOpsType(Chartertype, ICAO)=="Yes")
		            				{
			            				if(Impact=="A") twenty4H_APClsdCharters+="<tr><th bgcolor=\"Yellow\" style=\"color:SeaGreen;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="R") twenty4H_RWYClsdCharters+="<tr><th bgcolor=\"Yellow\" style=\"color:SeaGreen;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="C") twenty4H_CatICharters+="<tr><th bgcolor=\"Yellow\" style=\"color:SeaGreen;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="N") twenty4H_NilsCharters+="<tr><th bgcolor=\"Yellow\" style=\"color:SeaGreen;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="D") twenty4H_NoAltnCharters+="<tr><th bgcolor=\"Yellow\" style=\"color:SeaGreen;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="F") twenty4H_FuelCharters+="<tr><th bgcolor=\"Yellow\" style=\"color:SeaGreen;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="M") twenty4H_MiscCharters+="<tr><th bgcolor=\"Yellow\" style=\"color:SeaGreen;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th bgcolor=\"Yellow\" style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
		            				}

            					}
            					
		            			//Beyond 24 hrs
		            			//------------
		            			else
            					{
		            				if(!dBreader.IsDBNull(8)) ICAO = dBreader.GetString(8);
		            				if(!dBreader.IsDBNull(7)) all = dBreader.GetString(7);
		            				if(!dBreader.IsDBNull(10)) key = dBreader.GetString(10);
		            				if(!dBreader.IsDBNull(13)) Impact = dBreader.GetString(13);
		            				if(!dBreader.IsDBNull(14)) Remark = dBreader.GetString(14);
		            				
		            				//location = "<span style=\"color: blue;font-weight:bold\">"+ ICAO +" - "+GetIATA(ICAO) + " : </span>";
		            				location = GetIATA(ICAO);
		            				
		            				fromDate = dateTransformation(fromDate);
		            				tillDate = dateTransformation(tillDate);
		            				
		            				
		            				if(IsOpsType(LHtype, ICAO)=="Yes")
		            				{
			            				if(Impact=="A") APClsdLH+="<tr><th style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="R") RWYClsdLH+="<tr><th style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="C") CatILH+="<tr><th style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="N") NilsLH+="<tr><th style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="D") NoAltnLH+="<tr><th style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="F") FuelLH+="<tr><th style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="M") MiscLH+="<tr><th style=\"color:RoyalBlue;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
		            				}
		            				
		            				if(IsOpsType(SHtype, ICAO)=="Yes")
		            				{
			            				if(Impact=="A") APClsdSH+="<tr><th style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="R") RWYClsdSH+="<tr><th style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="C") CatISH+="<tr><th style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="N") NilsSH+="<tr><th style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="D") NoAltnSH+="<tr><th style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="F") FuelSH+="<tr><th style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="M") MiscSH+="<tr><th style=\"color:RebeccaPurple;\">"+location+"</th><th style=\"font-family:Courier New;\">"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
		            				}
		            				
		            				if(IsOpsType(Chartertype, ICAO)=="Yes")
		            				{
			            				if(Impact=="A") APClsdCharters+="<tr><th style=\"color:SeaGreen;\">"+location+"</th><th>"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="R") RWYClsdCharters+="<tr><th style=\"color:SeaGreen;\">"+location+"</th><th>"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="C") CatICharters+="<tr><th style=\"color:SeaGreen;\">"+location+"</th><th>"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="N") NilsCharters+="<tr><th style=\"color:SeaGreen;\">"+location+"</th><th>"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="D") NoAltnCharters+="<tr><th style=\"color:SeaGreen;\">"+location+"</th><th>"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="F") FuelCharters+="<tr><th style=\"color:SeaGreen;\">"+location+"</th><th>"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
			            				if(Impact=="M") MiscCharters+="<tr><th style=\"color:SeaGreen;\">"+location+"</th><th>"+key+"</th><th style=\"font-family:Courier New;\">"+fromDate+"-"+tillDate+"</th><th style=\"width:400px; font-weight:normal;\">"+Remark+"</th></tr>";
		            				}
            					}
		            			//RchTxtCSV.Text=twenty4H_APClsdCharters;
            				}
            			}
            		}
            	}
        	}
     		RchTxt_FilterNotams.Text=output;
     		
     		// *********************
     		string window="";
     		if(radBtn_noFilter.Checked)window="COMPLETE";
     		if(radBtn_24Hrs.Checked)window="Next 24 Hours";
     		if(radBtn_7days.Checked)window="Next 7 days";
     		if(radBtn_31days.Checked)window="Next 31 days";
     		
     		string reportDate= DateTime.Now.ToString("ddMMMMyyyy HHmm")+"CET";
     		
     		report="<html><head><title>NOTAM REPORT</title><body style=\"font-family:Calibri\">";
     		report+="<h1>Notam Report - "+window+"</h1>";
     		report+="<p>"+reportDate+"</p>";
     		report+="<table border=\"1\" style=\"width:700px; text-align: left; font-family:Calibri; font-size:12px; border : 1px solid black; border-collapse: collapse;\">";
     		report+="<tr><th colspan=\"4\" bgcolor=\"RoyalBlue\" style=\"font-size:16px; color: white; font-weight:bold; border : 1px solid black;\">" +
				"Long Haul :</th></tr><tr><th colspan=\"4\">"+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>AP Closed</b></th></tr>"+twenty4H_APClsdLH+APClsdLH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>RWY/TWY Closure impacting Perfos</b></th></tr>"+twenty4H_RWYClsdLH+RWYClsdLH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>CAT I</b></th></tr>"+twenty4H_CatILH+CatILH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>No ILS</b></th></tr>"+twenty4H_NilsLH+NilsLH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Not as Altn</b></th></tr>"+twenty4H_NoAltnLH+NoAltnLH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Fuel</b></th></tr>"+twenty4H_FuelLH+FuelLH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Misc</b></th></tr>"+twenty4H_MiscLH+MiscLH;
     		
     		report+="<tr><th colspan=\"4\" bgcolor=\"RebeccaPurple\" style=\"font-size:16px; color: white; font-weight:bold; border : 1px solid black;\">" +
				"Short Haul :</th></tr><tr><th colspan=\"4\">"+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>AP Closed</b></th></tr>"+twenty4H_APClsdSH+APClsdSH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>RWY/TWY Closure impacting Perfos</b></th></tr>"+twenty4H_RWYClsdSH+RWYClsdSH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>CAT I</b></th></tr>"+twenty4H_CatISH+CatISH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>No ILS</b></th></tr>"+twenty4H_NilsSH+NilsSH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Not as Altn</b></th></tr>"+twenty4H_NoAltnSH+NoAltnSH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Fuel</b></th></tr>"+twenty4H_FuelSH+FuelSH+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Misc</b></th></tr>"+twenty4H_MiscSH+MiscSH;
     		
     		report+="<tr><th colspan=\"4\" bgcolor=\"SeaGreen\" style=\"font-size:16px; color: white; font-weight:bold; border : 1px solid black;\">" +
				"Charters :</th></tr><tr><th colspan=\"4\">"+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>AP Closed</b></th></tr>"+twenty4H_APClsdCharters+APClsdCharters+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>RWY/TWY Closure impacting Perfos</b></th></tr>"+twenty4H_RWYClsdCharters+RWYClsdCharters+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>CAT I</b></th></tr>"+twenty4H_CatICharters+CatICharters+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>No ILS</b></th></tr>"+twenty4H_NilsCharters+NilsCharters+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Not as Altn</b></th></tr>"+twenty4H_NoAltnCharters+NoAltnCharters+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Fuel</b></th></tr>"+twenty4H_FuelCharters+FuelCharters+
     		"<tr><th colspan=\"4\" style=\"font-weight:bold;\" bgcolor=\"LightSlateGrey\" style=\"font-size:14px; color: white\"><b>Misc</b></th></tr>"+twenty4H_MiscCharters+MiscCharters;
     		
     			
     		report+="</table></body></html>";
     		
     		conn.Close();
     		
     		//On encode dans OCC
     		System.Data.OleDb.OleDbConnection conn2 = new System.Data.OleDb.OleDbConnection();
			conn2.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
				
			conn2.Open();
			
			var update = "UPDATE Notams_ICAO_CSV SET " +
		   	"APClsdLH='"+APClsdLH+"', " +
			"RWYClsdLH='"+RWYClsdLH+"', " +
			"CatILH='"+CatILH+"', " +	
			"NilsLH='"+NilsLH+"', " +
			"NoAltnLH='"+NoAltnLH+"', " +	
			"FuelLH='"+FuelLH+"', " +	
			"MiscLH='"+MiscLH+"', "+
			"APClsdSH='"+APClsdSH+"', " +
			"RWYClsdSH='"+RWYClsdSH+"', " +
			"CatISH='"+CatISH+"', " +	
			"NilsSH='"+NilsSH+"', " +
			"NoAltnSH='"+NoAltnSH+"', " +	
			"FuelSH='"+FuelSH+"', " +	
			"MiscSH='"+MiscSH+"', "+	
			"APClsdCharters='"+APClsdCharters+"', " +
			"RWYClsdCharters='"+RWYClsdCharters+"', " +
			"CatICharters='"+CatICharters+"', " +	
			"NilsCharters='"+NilsCharters+"', " +
			"NoAltnCharters='"+NoAltnCharters+"', " +	
			"FuelCharters='"+FuelCharters+"', " +	
			"MiscCharters='"+MiscCharters+"'"+	
           	"WHERE ID=1";
     		OleDbCommand commandUpdate = new OleDbCommand(update, conn2);
     		commandUpdate.ExecuteNonQuery();

			conn2.Close();
     		
			Web_report.DocumentText = report;
			
			// Base directory
	        string directoryPath = @"V:\TAY Ops Control Centre\Flight Dispatch\AIP SUP -  Notams report\Reports";
	
	        // Create reversed date filename (ddMMyy.html)
	        string fileName = DateTime.Now.ToString("yyMMdd") + "-Complete_NOTAMS_report.html";
	
	        // Full file path
	        string fullPath = Path.Combine(directoryPath, fileName);
	
	        // Ensure directory exists
	        Directory.CreateDirectory(directoryPath);
	
	        // Write file (overwrites if it already exists)
	        File.WriteAllText(fullPath, report);

			
		}
		// Cache en mémoire : ICAO → [IATA, LH, FedEx, Charters]
		// Chargé une seule fois au démarrage, rafraîchi après toute modification des stations.
		private static Dictionary<string, string[]> _stationsCache = null;

		public static void LoadStationsCache()
		{
			_stationsCache = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
			OleDbCommand cmd = new OleDbCommand("SELECT * FROM Stations_ICAO_IATA", conn);
			OleDbDataReader reader = cmd.ExecuteReader();
			while (reader.Read())
			{
				string icao = !reader.IsDBNull(1) ? reader.GetString(1) : "";
				if (icao == "") continue;
				string iata     = !reader.IsDBNull(2) ? reader.GetString(2) : "";
				string lh       = !reader.IsDBNull(3) ? reader.GetString(3) : "";
				string fedex    = !reader.IsDBNull(4) ? reader.GetString(4) : "";
				string charters = !reader.IsDBNull(5) ? reader.GetString(5) : "";
				_stationsCache[icao] = new string[] { iata, lh, fedex, charters };
			}
			conn.Close();
		}

		public static string IsOpsType(string OpsType, string location)
		{
			if (_stationsCache == null) LoadStationsCache();
			string[] row;
			if (!_stationsCache.TryGetValue(location, out row)) return "";
			if (OpsType == "LH")       return row[1];
			if (OpsType == "FedEx")    return row[2];
			if (OpsType == "Charters") return row[3];
			return "";
		}

		public static string GetIATA(string location)
		{
			if (_stationsCache == null) LoadStationsCache();
			string[] row;
			if (!_stationsCache.TryGetValue(location, out row)) return "";
			string IATA = row[0];
     		return IATA;
		}
//		public void StartApp()
//		{
//			//string CSV="CSV";
//			//RchTxtCSV.Text = CSV;
//			
//		}
		
		public void Reload_text(){
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM filteredNotams_table";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
     		string testText="";
     		
     		string []keyList=new string[100000];
     		
     		int i=0;
     		if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(10)) keyList[i]= dBreader.GetString(10);
            
            		i++;
            	}
     		}
     		conn.Close();
     		
     		
     		string [,]notamsList= new string[2,100000] ;
     		
     		i=0;
     		
     		//notamsList[0,3]="Test 1";
     	    //notamsList[1,3]="Text";
     		string notamText="";
     		foreach (string key in keyList) {

     			notamsList[0,i]=key;
     			
     			
				notamsList[1,i]=notamText;

     			i++;
     			
     		}
     		
     		//for (int j=0; j<keyList.Length; j++) {
     			int j=440;
     			string keyTest=keyList[j];
	     	    conn.Open();
	 			//var query3 = "SELECT * FROM storedNotams_table WHERE key='"+keyTest+"'";
	 			var query3 = "SELECT * FROM storedNotams_table WHERE ID=46511";
	 			OleDbCommand command5 = new OleDbCommand(query3, conn);
	 			OleDbDataReader dBreader2 = command5.ExecuteReader();
	 			
	 			if (dBreader2.HasRows)
	        	{
	            	while (dBreader2.Read())
	            	{
	            		if(!dBreader2.IsDBNull(15)) notamsList[1,j]= dBreader2.GetString(15);
	            
	            
	            	}
	     		}
	     		conn.Close();
     		//}
     	    
     		testText=notamsList[0,440]+" "+keyTest+" "+notamsList[1,440];
     		
     		RchTxtCSV.Text=testText;
     		
     		
     		
		}
	
		
		public void GetXML(){
			
			//Airport List from OCC.mdb
			string APList="";
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM Stations_ICAO_IATA";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
			
			//SqlCommand command = new SqlCommand("SELECT ICAO FROM Stations_ICAO_IATA",conn);			
			
        	//SqlDataReader dBreader = command.ExecuteReader();

        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		
                	if(!dBreader.IsDBNull(0)) 
                	{
                		string LH=dBreader.GetString(3);
                		string FedEx=dBreader.GetString(4);
                		string Charters=dBreader.GetString(5);
                		if(LH=="Yes"||FedEx=="Yes"||Charters=="Yes")
                		{
                			APList+=dBreader.GetString(1)+"-";
                		}
                	}
            	}
        	}
			
			conn.Close();
			APList = APList.Substring(0,APList.Length -1);
			
			var deletelog = "DELETE FROM storedNotams_table";
					
			System.Data.OleDb.OleDbConnection connNotams = new System.Data.OleDb.OleDbConnection();
			connNotams.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
					
			connNotams.Open();
			OleDbCommand commandedelete = new OleDbCommand(deletelog, connNotams);
			commandedelete.ExecuteNonQuery();
			connNotams.Close();
			
			DateTime today = DateTime.Now;
        	DateTime threeMonths = today.AddDays(91);
            			
            var todayString= DateTime.Now.ToString("yyyyMMdd");
            
            string todayYear="";
            string todayMonth="";
            string todayDay="";
            
            todayYear=todayString.Substring(2,2);
     		todayMonth=todayString.Substring(4,2);
     		todayDay=todayString.Substring(6,2);
     		
     		if(todayMonth=="01")todayMonth="JAN";
     		if(todayMonth=="02")todayMonth="FEB";
     		if(todayMonth=="03")todayMonth="MAR";
     		if(todayMonth=="04")todayMonth="APR";
     		if(todayMonth=="05")todayMonth="MAY";
     		if(todayMonth=="06")todayMonth="JUN";
     		if(todayMonth=="07")todayMonth="JUL";
     		if(todayMonth=="08")todayMonth="AUG";
     		if(todayMonth=="09")todayMonth="SEP";
     		if(todayMonth=="10")todayMonth="OCT";
     		if(todayMonth=="11")todayMonth="NOV";
     		if(todayMonth=="12")todayMonth="DEC";
     		
     		string stringToday="";
     		stringToday=todayDay+todayMonth+todayYear;
            
            var threeMonthsString= threeMonths.ToString("yyyyMMdd");
            
            string threeMonthsYear="";
            string threeMonthsMonth="";
            string threeMonthsDay="";
            
            threeMonthsYear=threeMonthsString.Substring(2,2);
     		threeMonthsMonth=threeMonthsString.Substring(4,2);
     		threeMonthsDay=threeMonthsString.Substring(6,2);
     		
     		if(threeMonthsMonth=="01")threeMonthsMonth="JAN";
     		if(threeMonthsMonth=="02")threeMonthsMonth="FEB";
     		if(threeMonthsMonth=="03")threeMonthsMonth="MAR";
     		if(threeMonthsMonth=="04")threeMonthsMonth="APR";
     		if(threeMonthsMonth=="05")threeMonthsMonth="MAY";
     		if(threeMonthsMonth=="06")threeMonthsMonth="JUN";
     		if(threeMonthsMonth=="07")threeMonthsMonth="JUL";
     		if(threeMonthsMonth=="08")threeMonthsMonth="AUG";
     		if(threeMonthsMonth=="09")threeMonthsMonth="SEP";
     		if(threeMonthsMonth=="10")threeMonthsMonth="OCT";
     		if(threeMonthsMonth=="11")threeMonthsMonth="NOV";
     		if(threeMonthsMonth=="12")threeMonthsMonth="DEC";
     		
     		string stringThreeMonths="";
     		stringThreeMonths=threeMonthsDay+threeMonthsMonth+threeMonthsYear;
            
     		RchTxt_FilterNotams.Text=stringToday+" "+stringThreeMonths;
                        				
            //RchTxt_FilterNotams.Text=todayYear+" "+threeMonthsYear;
			
			string xmlNotams="";
        	using (WebClient wc = new WebClient())
        	{
        		xmlNotams = wc.DownloadString("http://10.48.12.43:5455/BriefingService.svc/web/?METHOD=getAdHocNOTAM&AIRPORTS="+APList+"&PERIODSTART="+stringToday+"&PERIODEND="+stringThreeMonths+"");
        	}
        	
        	xmlNotams = Regex.Replace(xmlNotams, @"\t|\n|\r", " ");
        	
        	string[] splitOnNotamIssuer = Regex.Split(xmlNotams, "<NOTAM issuer=");
        	
        	string key="";
        	string series="";
        	string serial="";
        	string year="";
        	string all="";
        	string result="";
        	string startDate="";
        	string endDate="";
        	string creationDate="";
        	string revisionDate="";
        	string location="";
        	string NOTAMInfo="";
        	connNotams.Open();
        	
        	foreach (string search1 in splitOnNotamIssuer)
        	{
        		string[] splitOnSeries = Regex.Split(search1, "series=\"");
        		foreach (string search2 in splitOnSeries)
        		{
        			series=search2.Substring(0,1);
        			if(series!="<" && series!="\"" && series!="0") 
        			{
        				string[] splitOnSerial = Regex.Split(search2, "serial=\"");
        				foreach (string search3 in splitOnSerial)
        				{
        					if(search3.Length>4)
        					{
        						serial=search3.Substring(0,4);
        						string[] splitOnYear = Regex.Split(search3, "year=\"");
		        				foreach (string search4 in splitOnYear)
		        				{
		        					if(search4.Length>2)
		        					{
		        						year=search4.Substring(0,2);
		        						
		        						string[] splitOnStart = Regex.Split(search4, "startValidTime=\"");
				        				foreach (string search5 in splitOnStart)
				        				{
				        					if(search5.Length>19)
				        					{
				        						startDate=search5.Substring(0,19)+".000Z";
				        						string[] splitOnEnd = Regex.Split(search5, "endValidTime=\"");
									        	foreach (string search6 in splitOnEnd)
									        	{
									        		if(search6.Length>19)
									        		{
									        			endDate=search6.Substring(0,19)+".000Z";
									        			string[] splitOnCreation= Regex.Split(search6, "creationTime=\"");
											        	foreach (string search7 in splitOnCreation)
											        	{
											        		if(search7.Length>19)
											        		{
											        			creationDate=search7.Substring(0,19)+".000Z";
											        			string[] splitOnRevision= Regex.Split(search7, "revisionTime=\"");
													        	foreach (string search8 in splitOnCreation)
													        	{
													        		if(search8.Length>19)
													        		{
											        					revisionDate=search8.Substring(0,19)+".000Z";
										        						string[] splitOnText = Regex.Split(search8, "<NOTAMText><Paragraph><Text>");
													        			foreach (string search9 in splitOnText)
													        			{
													        				if(search9.Length>2)
													        				{
													        					string[] splitOnTextEnd = Regex.Split(search9, "</Text>");
													        					all=splitOnTextEnd[0];
													        					
													        					string[] splitOnICAO= Regex.Split(search9, "<AirportICAOCode>");
													        					foreach (string search10 in splitOnICAO)
															        			{
															        				if(search10.Length>4)
															        				{
															        					location=search10.Substring(0,4);
															        					
															        					NOTAMInfo="No";
															        					if(Regex.IsMatch(search10, "</ItemC><ItemD>"))
															        					{   
															        						NOTAMInfo="Yes";
																        					string[] splitOnNOTAMInfo= Regex.Split(search10, "</ItemC><ItemD>");
																        					foreach (string search11 in splitOnNOTAMInfo)
																		        			{
																		        				if(search11.Length>4)
																		        				{
																		        					string[] splitOnNOTAMInfoEnd= Regex.Split(search11, "</ItemD>");
																		        					NOTAMInfo=splitOnNOTAMInfoEnd[0];
																		           				} 	
																        					}
															        					}
															        				}
													        					}
													        				}
													        			}
													        		}
													        	}
											        		}
											        	}
									        		}
									        	}
				        					}
				        				}
		           					} 					
		        				}
           					} 					
        				}
        				
        				key=series+serial+"/"+year+"-"+location;
        				result+= key+"\n"+startDate+" - "+endDate+"\n"+all+"\nCreated: "+creationDate+"\nRevised:"+revisionDate+"\n"+"\n";
        				
        				if(NOTAMInfo!="") all= NOTAMInfo+"\n"+all;
        				all=all.Replace("\'","(char)39");
                		all=all.Replace("\"","(char)34");
        				
        				var insertlog = "INSERT INTO storedNotams_table ([Notam_id], [startdate], [enddate], [all], [location], [Created], [key])"
					    +" VALUES (?,?,?,?,?,?,?)";
			     	
//						try
//						{
							//connNotams.Open();
							OleDbCommand commandeinsert = new OleDbCommand(insertlog, connNotams);
							commandeinsert.Parameters.AddWithValue("?", key);
							commandeinsert.Parameters.AddWithValue("?", startDate);
							commandeinsert.Parameters.AddWithValue("?", endDate);
							commandeinsert.Parameters.AddWithValue("?", all);
							commandeinsert.Parameters.AddWithValue("?", location);
							commandeinsert.Parameters.AddWithValue("?", creationDate+" // "+revisionDate);
							commandeinsert.Parameters.AddWithValue("?", key);
							// Execution
							commandeinsert.ExecuteNonQuery();
							//connNotams.Close();
	                    }
        				
        				
        			//}
        			
        		}
            	
    		

        	}
        	connNotams.Close();
        	
        	RchTxtCSV.Text=result;
		}
		
		public void GetCSV()
		{
			//Airport List from OCC.mdb
			string APList="";
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
				
			conn.Open();
			var query2 = "SELECT * FROM Stations_ICAO_IATA";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
			
			//SqlCommand command = new SqlCommand("SELECT ICAO FROM Stations_ICAO_IATA",conn);			
			
        	//SqlDataReader dBreader = command.ExecuteReader();

        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		
                	if(!dBreader.IsDBNull(0)) 
                	{
                		string LH=dBreader.GetString(3);
                		string FedEx=dBreader.GetString(4);
                		string Charters=dBreader.GetString(5);
                		if(LH=="Yes"||FedEx=="Yes"||Charters=="Yes")
                		{
                			APList+=dBreader.GetString(1)+",";
                		}
                	}
            	}
        	}
			
			conn.Close();
			APList = APList.Substring(0,APList.Length -1);
			
			//c62b1e08-c60e-41ba-a654-30cb2807a682
			
			//string url = @"https://applications.icao.int/dataservices/api/notams-list?api_key=58af81e0-76b4-11eb-a919-8bb7115a8cf6&format=csv&type=&Qcode=&locations="+APList+"&qstring=&states=&ICAOonly=";
			string url = @"https://applications.icao.int/dataservices/api/notams-list?api_key=c62b1e08-c60e-41ba-a654-30cb2807a682&format=csv&type=&Qcode=&locations="+APList+"&qstring=&states=&ICAOonly=";
			
				
			//SSL/TLS config (works for HttpWebResponse & WebClient methods)
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
			
			//DOWNLOAD SYNC, OK for the sequence
			using (WebClient webClient = new WebClient())
			{
    			//webClient.DownloadFile("https://applications.icao.int/dataservices/api/notams-list?api_key=58af81e0-76b4-11eb-a919-8bb7115a8cf6&format=csv&type=&Qcode=&locations="+APList+"&qstring=&states=&ICAOonly=", "ICAO-CSV.csv");
				webClient.DownloadFile("https://applications.icao.int/dataservices/api/notams-list?api_key=c62b1e08-c60e-41ba-a654-30cb2807a682&format=csv&type=&Qcode=&locations="+APList+"&qstring=&states=&ICAOonly=", "ICAO-CSV.csv");
			
			}
			
			RchTxtCSV.Text="CSV Downloaded";
			
			//DONWLOAD ANSYNC, not OK as we want to operate in sequence
			
//    		using (WebClient wc = new WebClient())
//    		{
//        		wc.DownloadProgressChanged += wc_DownloadProgressChanged;
//        		wc.DownloadFileAsync (
//            	// Param1 = Link of file
//            	new System.Uri("https://applications.icao.int/dataservices/api/notams-list?api_key=58af81e0-76b4-11eb-a919-8bb7115a8cf6&format=csv&type=&Qcode=&locations="+APList+"&qstring=&states=&ICAOonly="),
//            	//new System.Uri(url),
//            	
//            	// Param2 = Path to save
//            	"ICAO-CSV.csv");
//    		}
    	
		}
		void Btn_CSVClick(object sender, EventArgs e)
		{
			GetCSV();
		}
		
	
		// Event to track the progress
		void wc_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
		{
    		progressBar.Value = e.ProgressPercentage;
		}
		
		public void Split()
		{
			var deletelog = "DELETE FROM storedNotams_table";
					
			System.Data.OleDb.OleDbConnection connNotams = new System.Data.OleDb.OleDbConnection();
			connNotams.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
					
			connNotams.Open();
			OleDbCommand commandedelete = new OleDbCommand(deletelog, connNotams);
			commandedelete.ExecuteNonQuery();
			connNotams.Close();
			
			var lines = File.ReadAllLines("ICAO-CSV.csv");
			string reader ="";
            int dataRowStart = 0;
            for (int i = dataRowStart; i < lines.Length; i++)
            {
                if(i==0)reader+= lines[i]+"\n";
                if(i>0)
                {	
                	Regex CSVParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
                	String[] subs = CSVParser.Split(lines[i]);
                	
                	string stringStateName="";
                	string stringStateCode="";
                	string stringid="";
                	string stringentity="";
                	string stringstatus="";
                	string stringQcode="";
                	string stringArea="";
                	string stringSubArea="";
                	string stringCondition="";
                	string stringSubject="";
                	string stringModifier="";
                	string stringmessage="";
                	string stringstartdate="";
                	string stringenddate="";
                	string stringall="";
                	string stringlocation="";
                	string stringisICAO="";
                	string stringCreated="";
                	string stringkey="";
                	string stringtype="";
                	
                	for(int j=0; j<subs.Length; j++)
                	{
                		if(subs[j].Length>2)subs[j]=subs[j].Substring(1,subs[j].Length -2);
                		//US start date and endate have 3 X double quotes
                		if(subs[j].Length>2)if(subs[j].Substring(0,2)=="\"\"") subs[j]=subs[j].Substring(2,subs[j].Length -4);
                		subs[j]=subs[j].Replace("\'","(char)39");
                		subs[j]=subs[j].Replace("\"","(char)34");
                		
                		if(j==0)  
                		{
                			stringStateName=subs[j];
                			reader+=stringStateName+"\n";
                		}
                		if(j==1)  
                		{
                			stringStateCode=subs[j];
                			reader+=stringStateCode+"\n";
                		}
                		if(j==2)  
                		{
                			stringid=subs[j];
                			reader+=stringid+"\n";
                		}
                		if(j==3)  
                		{
                			stringentity=subs[j];
                			reader+=stringentity+"\n";
                		}
                		if(j==4)  
                		{
                			stringstatus=subs[j];
                			reader+=stringstatus+"\n";
                		}
                		if(j==5)  
                		{
                			stringQcode=subs[j];
                			reader+=stringQcode+"\n";
                		}
                		if(j==6)  
                		{
                			stringArea=subs[j];
                			reader+=stringArea+"\n";
                		}
                		if(j==7)  
                		{
                			stringSubArea=subs[j];
                			reader+=stringSubArea+"\n";
                		}
                		if(j==8)  
                		{
                			stringCondition=subs[j];
                			reader+=stringCondition+"\n";
                		}
                		if(j==9)  
                		{
                			stringSubject=subs[j];
                			reader+=stringSubject+"\n";
                		}
                		if(j==10) 
                		{
                			stringModifier=subs[j];
                			reader+=stringModifier+"\n";
                		}
                		if(j==11) 
                		{
                			stringmessage=subs[j];
                			reader+=stringmessage+"\n";
                		}
                		if(j==12) 
                		{
                			stringstartdate=subs[j];
                			reader+=stringstartdate+"\n";
                		}
                		if(j==13) 
                		{
                			stringenddate=subs[j];
                			reader+=stringenddate+"\n";
                		}
                		if(j==14) 
                		{
                			stringall=subs[j];
                			reader+=stringall+"\n";
                		}
                		if(j==15) 
                		{
                			stringlocation=subs[j];
                			reader+=stringlocation+"\n";
                		}
                		if(j==16) 
                		{
                			stringisICAO=subs[j];
                			reader+=stringisICAO+"\n";
                		}
                		if(j==17) 
                		{
                			stringCreated=subs[j];
                			reader+=stringCreated+"\n";
                		}
                		if(j==18) 
                		{
                			stringkey=subs[j];
                			reader+=stringkey+"\n";
                		}
                		if(j==19) 
                		{
                			stringtype=subs[j];
                			reader+=subs[j]+"\n \n ------------------------------- \n \n";
                		}
                		
                	}
                	
                	var insertlog = "INSERT INTO storedNotams_table ([StateName], [StateCode], [Notam_id], [entity], [status], [Qcode], [Area], [SubArea], [Condition], [Subject], [Modifier], [message], [startdate], [enddate], [all], [location], [isICAO], [Created], [key], [type])"
					+" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

//					try
//					{
						connNotams.Open();
						OleDbCommand commandeinsert = new OleDbCommand(insertlog, connNotams);
						commandeinsert.Parameters.AddWithValue("?", stringStateName);
						commandeinsert.Parameters.AddWithValue("?", stringStateCode);
						commandeinsert.Parameters.AddWithValue("?", stringid);
						commandeinsert.Parameters.AddWithValue("?", stringentity);
						commandeinsert.Parameters.AddWithValue("?", stringstatus);
						commandeinsert.Parameters.AddWithValue("?", stringQcode);
						commandeinsert.Parameters.AddWithValue("?", stringArea);
						commandeinsert.Parameters.AddWithValue("?", stringSubArea);
						commandeinsert.Parameters.AddWithValue("?", stringCondition);
						commandeinsert.Parameters.AddWithValue("?", stringSubject);
						commandeinsert.Parameters.AddWithValue("?", stringModifier);
						commandeinsert.Parameters.AddWithValue("?", stringmessage);
						commandeinsert.Parameters.AddWithValue("?", stringstartdate);
						commandeinsert.Parameters.AddWithValue("?", stringenddate);
						commandeinsert.Parameters.AddWithValue("?", stringall);
						commandeinsert.Parameters.AddWithValue("?", stringlocation);
						commandeinsert.Parameters.AddWithValue("?", stringisICAO);
						commandeinsert.Parameters.AddWithValue("?", stringCreated);
						commandeinsert.Parameters.AddWithValue("?", stringkey);
						commandeinsert.Parameters.AddWithValue("?", stringtype);
						// Execution
						commandeinsert.ExecuteNonQuery();
						connNotams.Close();
                	
                }
            }
            reader=reader.Replace("(char)39","\'");
            reader=reader.Replace("(char)34","\"");
            RchTxtCSV.Text=reader;
            //"StateName","StateCode","id","entity","status","Qcode","Area","SubArea","Condition","Subject","Modifier","message","startdate","enddate","all","location","isICAO","Created","key","type"    
		}
		void Btn_splitClick(object sender, EventArgs e)
		{
			Split();
		}
		
		public void deleteWithdrawnedNotams()
		{
			System.Data.OleDb.OleDbConnection connNotams = new System.Data.OleDb.OleDbConnection();
			connNotams.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
			connNotams.Open();
			var query2 = "SELECT key FROM filteredNotams_table";
			OleDbCommand command4 = new OleDbCommand(query2, connNotams);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
     		string []notamFiltered= new string[100000];
     		
     		int i=0;
        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(0)) notamFiltered[i]=dBreader.GetString(0);
                	i++;    
            	}
        	}
        	connNotams.Close();
        	
        	
        	connNotams.Open();
        	//CHECK IF NOTAM IS Withdrawned AND RECORD in ARRAY:
        	string []notamWithdrawnedNOTAM= new string[100000];
        	for(int j =0;j<notamFiltered.Length;j++)
        	{
        		if(notamFiltered[j]!=null)
        		{  		
        			var queryCheckIfWithdrawned = "SELECT COUNT(*) FROM storedNotams_table WHERE key=?";
        			OleDbCommand commandCheckIfWithdrawned = new OleDbCommand(queryCheckIfWithdrawned, connNotams);
        			commandCheckIfWithdrawned.Parameters.AddWithValue("?", notamFiltered[j]);
     				Int32 count = (Int32) commandCheckIfWithdrawned.ExecuteScalar();
     			
     				if(count==0)notamWithdrawnedNOTAM[j]+=notamFiltered[j];
        		}
        	}
        	connNotams.Close();
        	
        	//Delete the withdrawned NOTAMS from filteredNOTAM$
        	string notamDelMonitor="";
        	connNotams.Open();
        	for(int k =0;k<notamWithdrawnedNOTAM.Length;k++)
        	{
				if(notamWithdrawnedNOTAM[k]!=null)
        		{
		        	var deletelog = "DELETE FROM filteredNotams_table WHERE key=?";

					OleDbCommand commandedelete = new OleDbCommand(deletelog, connNotams);
					commandedelete.Parameters.AddWithValue("?", notamWithdrawnedNOTAM[k]);
					commandedelete.ExecuteNonQuery();
					
					notamDelMonitor+=notamWithdrawnedNOTAM[k]+"\n";
				}
        	}
        	connNotams.Close();
        	RchTxtCSV.Text=notamDelMonitor;
		}
		
		public void NewNotams()
		{
			System.Data.OleDb.OleDbConnection connNotams = new System.Data.OleDb.OleDbConnection();
			connNotams.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
					
			connNotams.Open();
			//string station = "EBLG";
			//var query2 = "SELECT key FROM storedNotams_table WHERE location='"+station+"'";
			var query2 = "SELECT key FROM storedNotams_table";
			OleDbCommand command4 = new OleDbCommand(query2, connNotams);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
     		//CHECK THE NOTAMS ALREADY FILTERED AND STORE IN ARRAY
     		string testFilter="";
     		string []notamFiltered= new string[100000];
     		int i=0;
        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		string notamKey ="";
            		string all = "";
                	//if(!dBreader.IsDBNull(0)) notamKey=dBreader.GetString(0);
                	if(!dBreader.IsDBNull(0)) notamFiltered[i]=dBreader.GetString(0);
                	i++;    
            	}
        	}
        	connNotams.Close();
        	connNotams.Open();
        	//CHECK IF NOTAM IS NEW AND RECORD in ARRAY:
        	string []notamNewNOTAM= new string[100000];
        	for(int j =0;j<notamFiltered.Length;j++)
        	{
        		if(notamFiltered[j]!=null)
        		{  		
        			var queryCheckIfNew = "SELECT COUNT(*) FROM filteredNotams_table WHERE key=?";
        			OleDbCommand commandCheckIfNew = new OleDbCommand(queryCheckIfNew, connNotams);
        			commandCheckIfNew.Parameters.AddWithValue("?", notamFiltered[j]);
     				Int32 count = (Int32) commandCheckIfNew.ExecuteScalar();
     			
     				if(count==0)notamNewNOTAM[j]+=notamFiltered[j];
        		}
        	}
			connNotams.Close();
			for(int k =0;k<notamNewNOTAM.Length;k++)
        	{
				if(notamNewNOTAM[k]!=null)
        		{
					string StateName="";
					string Subject="";
					string Modifier="";
					string message="";
					string startdate="";
					string enddate="";
					string all="";
					string location="";
					string created="";
					string key=notamNewNOTAM[k];
					string Checked="N";
					connNotams.Open();
					var queryNewNotam = "SELECT * FROM storedNotams_table WHERE key=?";
					OleDbCommand commandNewNotam = new OleDbCommand(queryNewNotam, connNotams);
					commandNewNotam.Parameters.AddWithValue("?", notamNewNOTAM[k]);
     				OleDbDataReader NewNotamReader = commandNewNotam.ExecuteReader();					
					if (NewNotamReader.HasRows)
        			{
            			while (NewNotamReader.Read())
            			{
            				if(!NewNotamReader.IsDBNull(1)) StateName=NewNotamReader.GetString(1);
            				if(!NewNotamReader.IsDBNull(10)) Subject=NewNotamReader.GetString(10);
            				if(!NewNotamReader.IsDBNull(11)) Modifier=NewNotamReader.GetString(11);
            				if(!NewNotamReader.IsDBNull(12)) message=NewNotamReader.GetString(12);
            				if(!NewNotamReader.IsDBNull(13)) startdate=NewNotamReader.GetString(13);
            				if(!NewNotamReader.IsDBNull(14)) enddate=NewNotamReader.GetString(14);
            				if(!NewNotamReader.IsDBNull(15)) all=NewNotamReader.GetString(15);
            				if(!NewNotamReader.IsDBNull(16)) location=NewNotamReader.GetString(16);
            				if(!NewNotamReader.IsDBNull(18)) created=NewNotamReader.GetString(18);
            			}
					}
					
					connNotams.Close();
					//On encode le snouveaux notals ds FILTERED NOTAMS
					
					var insertNewNotam = "INSERT INTO filteredNotams_table ([StateName], [Subject], [Modifier], [message], [startdate], [enddate], [all], [location], [Created], [key], [Checked])"
					+" VALUES (?,?,?,?,?,?,?,?,?,?,?)";

					connNotams.Open();
					OleDbCommand commandeinsertNewNotam = new OleDbCommand(insertNewNotam, connNotams);
					commandeinsertNewNotam.Parameters.AddWithValue("?", StateName);
					commandeinsertNewNotam.Parameters.AddWithValue("?", Subject);
					commandeinsertNewNotam.Parameters.AddWithValue("?", Modifier);
					commandeinsertNewNotam.Parameters.AddWithValue("?", message);
					commandeinsertNewNotam.Parameters.AddWithValue("?", startdate);
					commandeinsertNewNotam.Parameters.AddWithValue("?", enddate);
					commandeinsertNewNotam.Parameters.AddWithValue("?", all);
					commandeinsertNewNotam.Parameters.AddWithValue("?", location);
					commandeinsertNewNotam.Parameters.AddWithValue("?", created);
					commandeinsertNewNotam.Parameters.AddWithValue("?", key);
					commandeinsertNewNotam.Parameters.AddWithValue("?", Checked);
					commandeinsertNewNotam.ExecuteNonQuery();
					connNotams.Close();
					
					testFilter+=StateName+"||"+Subject+"||"+Modifier+"||"+message+"||"+
						startdate+"||"+enddate+"||"+all+"||"+location+"||"+
						created+"||"+Checked+"\n";
					
					
				}
			}
			RchTxtCSV.Text=testFilter;
		}
		void Btn_newNotamsClick(object sender, EventArgs e)
		{
			NewNotams();
		}
		void Filter_Notams()
		{

			tabPage1.VerticalScroll.Value = 0;
			
			List<Label> itemsToRemove = new List<Label>();
			foreach (Label label in tabPage1.Controls.OfType<Label>())
			{
    			if (label.Tag != null && label.Tag.ToString() == "dispose")
    			{
      		  		itemsToRemove.Add(label);
    			}
			}
			foreach (Label label in itemsToRemove)
			{
    			tabPage1.Controls.Remove(label);
    			label.Dispose();
			}
			
			List<TextBox> txtboxToRemove = new List<TextBox>();
			foreach (TextBox txtbox in tabPage1.Controls.OfType<TextBox>())
			{
    			if (txtbox.Tag != null && txtbox.Tag.ToString() == "dispose")
    			{
      		  		txtboxToRemove.Add(txtbox);
    			}
			}
			foreach (TextBox txtbox in txtboxToRemove)
			{
    			tabPage1.Controls.Remove(txtbox);
    			txtbox.Dispose();
			}
			
			List<RichTextBox> rchtxtboxToRemove = new List<RichTextBox>();
			foreach (RichTextBox rchtxtbox in tabPage1.Controls.OfType<RichTextBox>())
			{
    			if (rchtxtbox.Tag != null && rchtxtbox.Tag.ToString() == "dispose")
    			{
      		  		rchtxtboxToRemove.Add(rchtxtbox);
    			}
			}
			foreach (RichTextBox rchtxtbox in rchtxtboxToRemove)
			{
    			tabPage1.Controls.Remove(rchtxtbox);
    			rchtxtbox.Dispose();
			}
			
			List<CheckBox> chckboxToRemove = new List<CheckBox>();
			foreach (CheckBox chckbox in tabPage1.Controls.OfType<CheckBox>())
			{
    			if (chckbox.Tag != null && chckbox.Tag.ToString() == "dispose")
    			{
      		  		chckboxToRemove.Add(chckbox);
    			}
			}
			foreach (CheckBox chckbox in chckboxToRemove)
			{
    			tabPage1.Controls.Remove(chckbox);
    			chckbox.Dispose();
			}
			
			List<Button> buttonsToRemove = new List<Button>();
			foreach (Button button in tabPage1.Controls.OfType<Button>())
			{
    			if (button.Tag != null && button.Tag.ToString() == "dispose")
    			{
      		  		buttonsToRemove.Add(button);
    			}
			}
			foreach (Button button in buttonsToRemove)
			{
    			tabPage1.Controls.Remove(button);
    			button.Dispose();
			}
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
			
			string AP="";
			conn.Open();
			var queryAP = "SELECT location FROM filteredNotams_table WHERE (Checked='N') ORDER BY location DESC";
     		OleDbCommand commandAP = new OleDbCommand(queryAP, conn);
     		OleDbDataReader APreader = commandAP.ExecuteReader();
     		
        	if (APreader.HasRows)
        	{
            	while (APreader.Read())
            	{
            		if(!APreader.IsDBNull(0)) AP=APreader.GetString(0);
            	}
        	}
			conn.Close();
			//OPEN OCC.MDB to get RWY's Infos
			System.Data.OleDb.OleDbConnection connOCC = new System.Data.OleDb.OleDbConnection();
			connOCC.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
						
			connOCC.Open();
			var queryOCC = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?";
     		OleDbCommand commandOCC = new OleDbCommand(queryOCC, connOCC);
     		commandOCC.Parameters.AddWithValue("?", AP);
     		OleDbDataReader OCCdBreader = commandOCC.ExecuteReader();
     		int AP_ID=0;
     		string RWYs="";
     		string richText="";
     					
     		Button[] update_Buttons = new Button[20000];
     		RichTextBox[] RchTxt_RWYs= new RichTextBox[20000];
			//Label[] ICAO= new Label[20000];
     		
        	if (OCCdBreader.HasRows)
        	{
            	while (OCCdBreader.Read())
            	{
            		if(!OCCdBreader.IsDBNull(0)) AP_ID=OCCdBreader.GetInt32(0);
            		if(!OCCdBreader.IsDBNull(6)) RWYs=OCCdBreader.GetString(6);
            	}
        	}
        	richText = AP+"\n";
        	
			//split of the RWY (NOT REQUIRED ANYMORE as entered from the form        	
        	string[] subRWYs = RWYs.Split('/');
        	foreach(string RWY in subRWYs)
        	{
        		richText+= RWY+"\n";
        	}
        	
        	richText+="_____________________________"+"\n"+"\n";
        	
        	connOCC.Close();
        	
        	//DIPLAY THE PREVIOUSLY KEPT NOTAMS in RICH TEXT FILTERNOTAM 
        	int notam_ID=0;
     		string notam_key="";
     		string notam_text="";
     		string location="";
     		string fromDate="";
     		string tillDate="";
     		
     		string fromHour="";
     		string fromDay="";
     		string fromMonth="";
     		string fromYear="";
     		string fromDateText="";
     		
     		string tillHour="";
     		string tillDay="";
     		string tillMonth="";
     		string tillYear="";
     		string tillDateText="";
     		
     		string Status="";
     		string Impact="";
     		string Remark="";
     		
        	conn.Open();
			var query1 = "SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)";
     		OleDbCommand command3 = new OleDbCommand(query1, conn);
     		command3.Parameters.AddWithValue("?", AP);
     		OleDbDataReader dBKeptreader = command3.ExecuteReader();
        	
     		if (dBKeptreader.HasRows)
        	{
            	while (dBKeptreader.Read())
            	{
            		if(!dBKeptreader.IsDBNull(0)) notam_ID=dBKeptreader.GetInt32(0);
            		if(!dBKeptreader.IsDBNull(5)) fromDate=dBKeptreader.GetString(5);
            		if(!dBKeptreader.IsDBNull(6)) tillDate=dBKeptreader.GetString(6);
            		if(!dBKeptreader.IsDBNull(7)) notam_text=dBKeptreader.GetString(7);
            		if(!dBKeptreader.IsDBNull(8)) location=dBKeptreader.GetString(8);
            		if(!dBKeptreader.IsDBNull(10)) notam_key=dBKeptreader.GetString(10);
            		if(!dBKeptreader.IsDBNull(12)) Status=dBKeptreader.GetString(12);
            		if(!dBKeptreader.IsDBNull(13)) Impact=dBKeptreader.GetString(13);
            		if(!dBKeptreader.IsDBNull(14)) Remark=dBKeptreader.GetString(14);
            		notam_text=notam_text.Replace("(char)39","'");
            		
            		fromDate=fromDate.Substring(0,16);
     				fromHour=fromDate.Substring(11,5);
     				fromDay=fromDate.Substring(8,2);
     				fromMonth=fromDate.Substring(5,2);
     				if(fromMonth=="01")fromMonth="JAN";
     				if(fromMonth=="02")fromMonth="FEB";
     				if(fromMonth=="03")fromMonth="MAR";
     				if(fromMonth=="04")fromMonth="APR";
     				if(fromMonth=="05")fromMonth="MAY";
     				if(fromMonth=="06")fromMonth="JUN";
     				if(fromMonth=="07")fromMonth="JUL";
     				if(fromMonth=="08")fromMonth="AUG";
     				if(fromMonth=="09")fromMonth="SEP";
     				if(fromMonth=="10")fromMonth="OCT";
     				if(fromMonth=="11")fromMonth="NOV";
     				if(fromMonth=="12")fromMonth="DEC";
     				fromYear=fromDate.Substring(0,4);
     				fromDateText=fromDay+fromMonth+fromYear+"("+fromHour+")";
     				
     				tillDate=tillDate.Substring(0,16);
     				tillHour=tillDate.Substring(11,5);
     				tillDay=tillDate.Substring(8,2);
     				tillMonth=tillDate.Substring(5,2);
     				if(tillMonth=="01")tillMonth="JAN";
     				if(tillMonth=="02")tillMonth="FEB";
     				if(tillMonth=="03")tillMonth="MAR";
     				if(tillMonth=="04")tillMonth="APR";
     				if(tillMonth=="05")tillMonth="MAY";
     				if(tillMonth=="06")tillMonth="JUN";
     				if(tillMonth=="07")tillMonth="JUL";
     				if(tillMonth=="08")tillMonth="AUG";
     				if(tillMonth=="09")tillMonth="SEP";
     				if(tillMonth=="10")tillMonth="OCT";
     				if(tillMonth=="11")tillMonth="NOV";
     				if(tillMonth=="12")tillMonth="DEC";
     				tillYear=tillDate.Substring(0,4);
     				tillDateText=tillDay+tillMonth+tillYear+"("+tillHour+")";
            		
            		richText+= notam_key+"\n"+fromDateText+" - "+tillDateText+"\n"+notam_text+"\n"+Impact+": "+Remark+"\n"+"\n";
            	}
     		}
     		conn.Close();
     		
     		RchTxt_FilterNotams.Text=richText;
     		
        	//DIPLAY THE NEW FILTERED NOTAMS
			conn.Open();
			var query2 = "SELECT * FROM filteredNotams_table WHERE (Checked='N') AND (location=?)";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		command4.Parameters.AddWithValue("?", AP);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		notam_ID=0;
     		notam_key="";
     		notam_text="";
     		location="";
     		fromDate="";
     		fromHour="";
     		fromDay="";
     		fromMonth="";
     		fromYear="";
     		fromDateText="";
     		tillDate="";
     		tillHour="";
     		tillDay="";
     		tillMonth="";
     		tillYear="";
     		tillDateText="";
     		int nbNotams = 0;
     		int Top = 0;
     					
     		Button[] keep_Buttons = new Button[20000];
     		RichTextBox[] RchTxt_notam_text= new RichTextBox[20000];
			CheckBox[] apt_CLSD_Chckbox= new CheckBox[20000];
			CheckBox[] apt_CATI_Chckbox= new CheckBox[20000];
			CheckBox[] apt_NILS_Chckbox= new CheckBox[20000];
			CheckBox[] apt_NOALTN_Chckbox= new CheckBox[20000];
			CheckBox[] apt_FUEL_Chckbox= new CheckBox[20000];
			CheckBox[] apt_MISC_Chckbox= new CheckBox[20000];
			CheckBox[] apt_AIPSUP_Chckbox= new CheckBox[20000];
			CheckBox[] apt_RWYCLSD_Chckbox= new CheckBox[20000];
			TextBox[] remark_Txtbox= new TextBox[20000];
			Button[] remark_Buttons = new Button[20000];
     		
			Status="";
     		Impact="";
     		Remark="";
     		
     		Top=100;
        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(0)) notam_ID=dBreader.GetInt32(0);
            		if(!dBreader.IsDBNull(5)) fromDate=dBreader.GetString(5);
            		if(!dBreader.IsDBNull(6)) tillDate=dBreader.GetString(6);
            		if(!dBreader.IsDBNull(7)) notam_text=dBreader.GetString(7);
            		if(!dBreader.IsDBNull(8)) location=dBreader.GetString(8);
            		if(!dBreader.IsDBNull(10)) notam_key=dBreader.GetString(10);
            		if(!dBreader.IsDBNull(12)) Status=dBreader.GetString(12);
            		if(!dBreader.IsDBNull(13)) Impact=dBreader.GetString(13);
            		if(!dBreader.IsDBNull(14)) Remark=dBreader.GetString(14);
            		notam_text=notam_text.Replace("(char)39","'");
            		
            		Label lbl_notam_key = new Label();
            		FontFamily courier =  new FontFamily("Courier New");
					lbl_notam_key.Font = new Font(courier, 10, FontStyle.Regular);
					lbl_notam_key.Tag ="dispose";
					lbl_notam_key.Top = Top;
     				lbl_notam_key.Size = new Size(125, 16);
     				lbl_notam_key.ForeColor = Color.Black;
     				lbl_notam_key.Text = notam_key;
     				lbl_notam_key.Left = 510;
     				tabPage1.Controls.Add(lbl_notam_key);
            		
     				fromDate=fromDate.Substring(0,16);
     				fromHour=fromDate.Substring(11,5);
     				fromDay=fromDate.Substring(8,2);
     				fromMonth=fromDate.Substring(5,2);
     				if(fromMonth=="01")fromMonth="JAN";
     				if(fromMonth=="02")fromMonth="FEB";
     				if(fromMonth=="03")fromMonth="MAR";
     				if(fromMonth=="04")fromMonth="APR";
     				if(fromMonth=="05")fromMonth="MAY";
     				if(fromMonth=="06")fromMonth="JUN";
     				if(fromMonth=="07")fromMonth="JUL";
     				if(fromMonth=="08")fromMonth="AUG";
     				if(fromMonth=="09")fromMonth="SEP";
     				if(fromMonth=="10")fromMonth="OCT";
     				if(fromMonth=="11")fromMonth="NOV";
     				if(fromMonth=="12")fromMonth="DEC";
     				fromYear=fromDate.Substring(0,4);
     				fromDateText=fromDay+fromMonth+fromYear+"("+fromHour+")";
     				Label lbl_notam_from = new Label();
            		//FontFamily courier =  new FontFamily("Courier New");
					lbl_notam_from.Font = new Font(courier, 10, FontStyle.Regular);
					lbl_notam_from.Tag ="dispose";
					lbl_notam_from.Top = Top;
     				lbl_notam_from.Size = new Size(200, 16);
     				lbl_notam_from.ForeColor = Color.Black;
     				lbl_notam_from.Text = "From : "+ fromDateText;
     				lbl_notam_from.Left = 635;
     				tabPage1.Controls.Add(lbl_notam_from);
     				
     				tillDate=tillDate.Substring(0,16);
     				tillHour=tillDate.Substring(11,5);
     				tillDay=tillDate.Substring(8,2);
     				tillMonth=tillDate.Substring(5,2);
     				if(tillMonth=="01")tillMonth="JAN";
     				if(tillMonth=="02")tillMonth="FEB";
     				if(tillMonth=="03")tillMonth="MAR";
     				if(tillMonth=="04")tillMonth="APR";
     				if(tillMonth=="05")tillMonth="MAY";
     				if(tillMonth=="06")tillMonth="JUN";
     				if(tillMonth=="07")tillMonth="JUL";
     				if(tillMonth=="08")tillMonth="AUG";
     				if(tillMonth=="09")tillMonth="SEP";
     				if(tillMonth=="10")tillMonth="OCT";
     				if(tillMonth=="11")tillMonth="NOV";
     				if(tillMonth=="12")tillMonth="DEC";
     				tillYear=tillDate.Substring(0,4);
     				tillDateText=tillDay+tillMonth+tillYear+"("+tillHour+")";
     				Label lbl_notam_till = new Label();
            		//FontFamily courier =  new FontFamily("Courier New");
					lbl_notam_till.Font = new Font(courier, 10, FontStyle.Regular);
					lbl_notam_till.Tag ="dispose";
					lbl_notam_till.Top = Top;
     				lbl_notam_till.Size = new Size(200, 16);
     				lbl_notam_till.ForeColor = Color.Black;
     				lbl_notam_till.Text = "Till : "+ tillDateText;
     				lbl_notam_till.Left = 840;
     				tabPage1.Controls.Add(lbl_notam_till);
     				
//            		int height = notam_text.Length/50*20;
//            		if(height<80)height=80;
//            		lbl_notam_text[notam_ID] = new Label();
//            		//FontFamily courier =  new FontFamily("Courier New");
//					lbl_notam_text[notam_ID].Font = new Font(courier, 10, FontStyle.Regular);
//					lbl_notam_text[notam_ID].Tag ="dispose";
//					lbl_notam_text[notam_ID].Top = Top+20;
//     				lbl_notam_text[notam_ID].Size = new Size(550, height);
//     				lbl_notam_text[notam_ID].ForeColor = Color.Black;
//     				lbl_notam_text[notam_ID].BackColor= Color.LightCoral;
//     				if(Status=="K")lbl_notam_text[notam_ID].BackColor= Color.CornflowerBlue;
//     				lbl_notam_text[notam_ID].Text = notam_text;
//     				lbl_notam_text[notam_ID].Left = 310;
//     				tabPage1.Controls.Add(lbl_notam_text[notam_ID] );
     				
     				int height = notam_text.Length/50*20;
            		if(height<80)height=80;
            		RchTxt_notam_text[notam_ID] = new RichTextBox();
            		//FontFamily courier =  new FontFamily("Courier New");
					RchTxt_notam_text[notam_ID].Font = new Font(courier, 10, FontStyle.Regular);
					RchTxt_notam_text[notam_ID].Tag ="dispose";
					RchTxt_notam_text[notam_ID].Top = Top+20;
     				RchTxt_notam_text[notam_ID].Size = new Size(550, height);
     				RchTxt_notam_text[notam_ID].ForeColor = Color.Black;
     				RchTxt_notam_text[notam_ID].BackColor= Color.LightCoral;
     				if(Status=="K")RchTxt_notam_text[notam_ID].BackColor= Color.CornflowerBlue;
     				RchTxt_notam_text[notam_ID].Text = notam_text;
     				RchTxt_notam_text[notam_ID].Left = 510;
     				RchTxt_notam_text[notam_ID].ReadOnly=true;
     				tabPage1.Controls.Add(RchTxt_notam_text[notam_ID] );
     						
     				keep_Buttons[notam_ID] = new Button();
     				keep_Buttons[notam_ID].Tag = "dispose";
     				keep_Buttons[notam_ID].Size = new Size(40 , 25);
     				if(Status=="K")keep_Buttons[notam_ID].Size = new Size(50 , 25);
        			keep_Buttons[notam_ID].Location = new Point(1070, Top+20);
        			int newSize =7;
        			keep_Buttons[notam_ID].Text = "Keep";
        			if(Status=="K")keep_Buttons[notam_ID].Text = "Ignore";
        			keep_Buttons[notam_ID].BackColor = Color.CornflowerBlue;
        			if(Status=="K")keep_Buttons[notam_ID].BackColor = Color.LightCoral;
        			keep_Buttons[notam_ID].Font = new Font(keep_Buttons[notam_ID].Font.FontFamily, newSize);
        			int i_notamToKeep = notam_ID;
        			if(Status=="")keep_Buttons[notam_ID].Click += (sender1, ex) => this.Keep_Notam(i_notamToKeep);
        			if(Status=="K")keep_Buttons[notam_ID].Click += (sender1, ex) => this.Ignore_Notam(i_notamToKeep);
     				tabPage1.Controls.Add(keep_Buttons[notam_ID]);	
     				
     				if(Status=="K")
     				{
     					apt_CLSD_Chckbox[notam_ID] = new CheckBox();
     					apt_CLSD_Chckbox[notam_ID].Tag ="dispose";
     					apt_CLSD_Chckbox[notam_ID].Top = Top+44;
     					apt_CLSD_Chckbox[notam_ID].Left = 1070;
     					apt_CLSD_Chckbox[notam_ID].Text = "APT CLSD";
     					apt_CLSD_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamAPTCLSD=notam_ID;
     					if(Impact=="A")apt_CLSD_Chckbox[notam_ID].Checked=true;
     					else apt_CLSD_Chckbox[notam_ID].Checked=false;
     					string A = "A";
     					apt_CLSD_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTCLSD, A);
     					tabPage1.Controls.Add(apt_CLSD_Chckbox[notam_ID]);
     					
     					apt_CATI_Chckbox[notam_ID] = new CheckBox();
     					apt_CATI_Chckbox[notam_ID].Tag ="dispose";
     					apt_CATI_Chckbox[notam_ID].Top = Top+44;
     					apt_CATI_Chckbox[notam_ID].Left = 1150;
     					apt_CATI_Chckbox[notam_ID].Text = "APT CATI";
     					apt_CATI_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamAPTCATI=notam_ID;
     					if(Impact=="C")apt_CATI_Chckbox[notam_ID].Checked=true;
     					else apt_CATI_Chckbox[notam_ID].Checked=false;
     					string C ="C";
     					apt_CATI_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTCATI, C);
     					tabPage1.Controls.Add(apt_CATI_Chckbox[notam_ID]);
     					
     					apt_NILS_Chckbox[notam_ID] = new CheckBox();
     					apt_NILS_Chckbox[notam_ID].Tag ="dispose";
     					apt_NILS_Chckbox[notam_ID].Top = Top+44;
     					apt_NILS_Chckbox[notam_ID].Left = 1240;
     					apt_NILS_Chckbox[notam_ID].Text = "No ILS";
     					apt_NILS_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamAPTNILS=notam_ID;
     					if(Impact=="N")apt_NILS_Chckbox[notam_ID].Checked=true;
     					else apt_NILS_Chckbox[notam_ID].Checked=false;
     					string N="N";
     					apt_NILS_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTNILS, N);
     					tabPage1.Controls.Add(apt_NILS_Chckbox[notam_ID]);
     					
     					apt_NOALTN_Chckbox[notam_ID] = new CheckBox();
     					apt_NOALTN_Chckbox[notam_ID].Tag ="dispose";
     					apt_NOALTN_Chckbox[notam_ID].Top = Top+68;
     					apt_NOALTN_Chckbox[notam_ID].Left = 1070;
     					apt_NOALTN_Chckbox[notam_ID].Text = "Not ALTN";
     					apt_NOALTN_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamAPTNOALTN=notam_ID;
     					if(Impact=="D")apt_NOALTN_Chckbox[notam_ID].Checked=true;
     					else apt_NOALTN_Chckbox[notam_ID].Checked=false;
     					string D="D";
     					apt_NOALTN_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTNOALTN, D);
     					tabPage1.Controls.Add(apt_NOALTN_Chckbox[notam_ID]);
     					
     					apt_FUEL_Chckbox[notam_ID] = new CheckBox();
     					apt_FUEL_Chckbox[notam_ID].Tag ="dispose";
     					apt_FUEL_Chckbox[notam_ID].Top = Top+68;
     					apt_FUEL_Chckbox[notam_ID].Left = 1150;
     					apt_FUEL_Chckbox[notam_ID].Text = "Fuel";
     					apt_FUEL_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamAPTFUEL=notam_ID;
     					if(Impact=="F")apt_FUEL_Chckbox[notam_ID].Checked=true;
     					else apt_FUEL_Chckbox[notam_ID].Checked=false;
     					string F="F";
     					apt_FUEL_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTFUEL, F);
     					tabPage1.Controls.Add(apt_FUEL_Chckbox[notam_ID]);
     					
     					apt_MISC_Chckbox[notam_ID] = new CheckBox();
     					apt_MISC_Chckbox[notam_ID].Tag ="dispose";
     					apt_MISC_Chckbox[notam_ID].Top = Top+68;
     					apt_MISC_Chckbox[notam_ID].Left = 1240;
     					apt_MISC_Chckbox[notam_ID].Text = "MISC";
     					apt_MISC_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamMISC=notam_ID;
     					if(Impact=="M")apt_MISC_Chckbox[notam_ID].Checked=true;
     					else apt_MISC_Chckbox[notam_ID].Checked=false;
     					string M="M";
     					apt_MISC_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamMISC, M);
     					tabPage1.Controls.Add(apt_MISC_Chckbox[notam_ID]);
     					
     					//******************************************
     					apt_AIPSUP_Chckbox[notam_ID] = new CheckBox();
     					apt_AIPSUP_Chckbox[notam_ID].Tag ="dispose";
     					apt_AIPSUP_Chckbox[notam_ID].Top = Top+44;
     					apt_AIPSUP_Chckbox[notam_ID].Left = 1330;
     					apt_AIPSUP_Chckbox[notam_ID].Text = "SUP";
     					apt_AIPSUP_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamAIPSUP=notam_ID;
     					if(Impact=="AS")apt_AIPSUP_Chckbox[notam_ID].Checked=true;
     					else apt_AIPSUP_Chckbox[notam_ID].Checked=false;
     					string AS="AS";
     					apt_AIPSUP_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAIPSUP, AS);
     					tabPage1.Controls.Add(apt_AIPSUP_Chckbox[notam_ID]);
     					
     					apt_RWYCLSD_Chckbox[notam_ID] = new CheckBox();
     					apt_RWYCLSD_Chckbox[notam_ID].Tag ="dispose";
     					apt_RWYCLSD_Chckbox[notam_ID].Top = Top+68;
     					apt_RWYCLSD_Chckbox[notam_ID].Left = 1330;
     					apt_RWYCLSD_Chckbox[notam_ID].Text = "RWY";
     					apt_RWYCLSD_Chckbox[notam_ID].Size = new Size (80,25);
     					int i_notamRWYCLSD=notam_ID;
     					if(Impact=="R")apt_RWYCLSD_Chckbox[notam_ID].Checked=true;
     					else apt_RWYCLSD_Chckbox[notam_ID].Checked=false;
     					string R="R";
     					apt_RWYCLSD_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamRWYCLSD, R);
     					tabPage1.Controls.Add(apt_RWYCLSD_Chckbox[notam_ID]);
     					
     					
     					if(Impact=="A" || Impact=="C" || Impact=="N"|| Impact=="D"|| Impact=="F"|| Impact=="M"|| Impact=="AS"|| Impact=="R")
     					{
     						remark_Txtbox[notam_ID] = new TextBox();
     						remark_Txtbox[notam_ID].Tag="dispose";
     						remark_Txtbox[notam_ID].Top = Top+94;
     						remark_Txtbox[notam_ID].Left = 1070;
     						remark_Txtbox[notam_ID].Size = new Size(250, 24);
     						remark_Txtbox[notam_ID].Text= Remark;
     						
     						remark_Buttons[notam_ID] = new Button();
     						remark_Buttons[notam_ID].Tag="dispose";
     						remark_Buttons[notam_ID].Top = Top+92;
     						remark_Buttons[notam_ID].Left = 1320;
     						remark_Buttons[notam_ID].Size = new Size(40, 24);
     						remark_Buttons[notam_ID].Text= "OK";
     						int i_remark=notam_ID;
     						
     						//remark_Txtbox[notam_ID].Leave += (sender1, ex) => this.Remark_Notam(i_remark, remark_Txtbox[i_remark].Text);
     						remark_Buttons[notam_ID].Click += (sender1, ex) => this.Remark_Notam(i_remark, remark_Txtbox[i_remark].Text);
     						tabPage1.Controls.Add(remark_Buttons[notam_ID]);
     						tabPage1.Controls.Add(remark_Txtbox[notam_ID]);
     					}
     				}
     				Top = Top+height+30;
     				nbNotams++;
     				Status="";
     				Impact="";
     				Remark="";
            	}       		
        	}
        	conn.Close();
            
            string notamsUnchecked="";
            notamsUnchecked="Notams Unchecked : "+ nbNotams;
            
            Lbl_location.Text=AP;
            Lbl_notamsUnchecked.Text=notamsUnchecked;
            //tabPage1.VerticalScroll.Value = 0;
            Btn_submitNotams.Top=Top+30;
		}
		void Keep_Notam(int notam_ID)
		{
			string stringIDNotam =notam_ID.ToString();
			//RchTxtCSV.Text=stringIDNotam;
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var update = "UPDATE filteredNotams_table SET Status='K' WHERE ID=?";
     		OleDbCommand command4 = new OleDbCommand(update, conn);
     		command4.Parameters.AddWithValue("?", notam_ID);
     		command4.ExecuteNonQuery();
     		conn.Close();
     		Filter_Notams();
     		ICAO_Notams();
		}
		void Ignore_Notam(int notam_ID)
		{
			string stringIDNotam =notam_ID.ToString();
			//RchTxtCSV.Text=stringIDNotam;
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var update = "UPDATE filteredNotams_table SET Status='' WHERE ID=?";
     		OleDbCommand command4 = new OleDbCommand(update, conn);
     		command4.Parameters.AddWithValue("?", notam_ID);
     		command4.ExecuteNonQuery();
     		conn.Close();
     		Filter_Notams();
     		ICAO_Notams();
		}
		public void DelOld()
		{
			string endDate="";
			string[] notamKeyList=new string[10000];
			string endDateList="";
			string todayList="";
			//string keyList="";
			int item=0;
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var query2 = "SELECT enddate,key FROM filteredNotams_table";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		DateTime _date = DateTime.Now;
     		_date=_date.ToUniversalTime();
     		
			string today = _date.ToString("yyyyMMddHHmm");
			todayList+=today+"\n";
			long intToday = Int64.Parse(today);
     		
        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(0)) endDate=dBreader.GetString(0);
            		endDate=endDate.Replace("-",string.Empty);
            		endDate=endDate.Replace("T",string.Empty);
            		endDate=endDate.Replace(":",string.Empty);
            		endDate=endDate.Replace(".",string.Empty);
            		endDate=endDate.Replace("Z",string.Empty);
            		long intEnd=0;
            		if(endDate.Length>11) 
            		{
            			endDate=endDate.Substring(0,12);
            			endDateList+= endDate+"\n";
            			intEnd= Int64.Parse(endDate);
            		}
            		
            		
       				
					if(intToday>intEnd)
					{
						endDateList+= endDate+"\n";
						if(!dBreader.IsDBNull(1)) notamKeyList[item]=dBreader.GetString(1);
						item++;
					}
					
            	}          		
        	}
        	conn.Close();
        	
        	string deletedNotams="";
        	
        	conn.Open();
        	for(int i=0;i<notamKeyList.Length;i++)
        	{
        		if(notamKeyList[i]!="")
        		{
        			var deletelog = "DELETE FROM filteredNotams_table WHERE key=?";

					OleDbCommand commandedelete = new OleDbCommand(deletelog, conn);
					commandedelete.Parameters.AddWithValue("?", notamKeyList[i]);
					commandedelete.ExecuteNonQuery();
					deletedNotams+=notamKeyList[i];
        		}
        	}
        	conn.Close();
        	
        	RchTxtCSV.Text=todayList+"\n\n"+deletedNotams+"\n\n"+endDateList;
		}
		void Btn_delOldClick(object sender, EventArgs e)
		{
			DelOld();
		}
		void Btn_analyzeNotamsClick(object sender, EventArgs e)
		{
			Filter_Notams();
		}
		void Btn_submitNotamsClick(object sender, EventArgs e)
		{
			string AP= Lbl_location.Text;
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
				
			conn.Open();
			var update = "UPDATE filteredNotams_table SET Checked='Y' WHERE (location=?) AND (Checked='N')";
     		OleDbCommand command4 = new OleDbCommand(update, conn);
     		command4.Parameters.AddWithValue("?", AP);
     		command4.ExecuteNonQuery();
     		conn.Close();
     		Filter_Notams();
		}
		void ICAO_Notams()
		{
			tabPage2.VerticalScroll.Value = 0;
			
			List<Label> itemsToRemove = new List<Label>();
			foreach (Label label in tabPage2.Controls.OfType<Label>())
			{
    			if (label.Tag != null && label.Tag.ToString() == "dispose")
    			{
      		  		itemsToRemove.Add(label);
    			}
			}
			foreach (Label label in itemsToRemove)
			{
    			tabPage2.Controls.Remove(label);
    			label.Dispose();
			}
			
			List<TextBox> txtboxToRemove = new List<TextBox>();
			foreach (TextBox txtbox in tabPage2.Controls.OfType<TextBox>())
			{
    			if (txtbox.Tag != null && txtbox.Tag.ToString() == "dispose")
    			{
      		  		txtboxToRemove.Add(txtbox);
    			}
			}
			foreach (TextBox txtbox in txtboxToRemove)
			{
    			tabPage2.Controls.Remove(txtbox);
    			txtbox.Dispose();
			}
			
			List<RichTextBox> rchtxtboxToRemove = new List<RichTextBox>();
			foreach (RichTextBox rchtxtbox in tabPage2.Controls.OfType<RichTextBox>())
			{
    			if (rchtxtbox.Tag != null && rchtxtbox.Tag.ToString() == "dispose")
    			{
      		  		rchtxtboxToRemove.Add(rchtxtbox);
    			}
			}
			foreach (RichTextBox rchtxtbox in rchtxtboxToRemove)
			{
    			tabPage2.Controls.Remove(rchtxtbox);
    			rchtxtbox.Dispose();
			}
			
			List<CheckBox> chckboxToRemove = new List<CheckBox>();
			foreach (CheckBox chckbox in tabPage2.Controls.OfType<CheckBox>())
			{
    			if (chckbox.Tag != null && chckbox.Tag.ToString() == "dispose")
    			{
      		  		chckboxToRemove.Add(chckbox);
    			}
			}
			foreach (CheckBox chckbox in chckboxToRemove)
			{
    			tabPage2.Controls.Remove(chckbox);
    			chckbox.Dispose();
			}
			
			List<Button> buttonsToRemove = new List<Button>();
			foreach (Button button in tabPage2.Controls.OfType<Button>())
			{
    			if (button.Tag != null && button.Tag.ToString() == "dispose")
    			{
      		  		buttonsToRemove.Add(button);
    			}
			}
			foreach (Button button in buttonsToRemove)
			{
    			tabPage2.Controls.Remove(button);
    			button.Dispose();
			}
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
			
			string AP=TxtBox_ICAO.Text;
			var query2="";
			//OPEN OCC.MDB to get RWY's Infos
			System.Data.OleDb.OleDbConnection connOCC = new System.Data.OleDb.OleDbConnection();
			connOCC.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
						
			connOCC.Open();
			var queryOCC = "SELECT * FROM Stations_ICAO_IATA WHERE ICAO=?";
     		OleDbCommand commandOCC = new OleDbCommand(queryOCC, connOCC);
     		commandOCC.Parameters.AddWithValue("?", AP);
     		OleDbDataReader OCCdBreader = commandOCC.ExecuteReader();
     		int AP_ID=0;
     		string RWYs="";

     		Button[] update_Buttons = new Button[20000];
     		RichTextBox[] RchTxt_RWYs= new RichTextBox[20000];
			//Label[] ICAO= new Label[20000];
     		
        	if (OCCdBreader.HasRows)
        	{
            	while (OCCdBreader.Read())
            	{
            		if(!OCCdBreader.IsDBNull(0)) AP_ID=OCCdBreader.GetInt32(0);
            		if(!OCCdBreader.IsDBNull(6)) RWYs=OCCdBreader.GetString(6);
            	}
        	}
        	string stationRWYs="";
        	stationRWYs=stationRWYs.Replace(" ","&nbsp");
        	stationRWYs=stationRWYs.Replace("\n","<br />");
        	stationRWYs="<p style=\"font:Courier New;\"><b><u>"+AP+"</u></b>"+ "<br />"+ RWYs+"</p>";
        	
        	Web_ICAONotams.DocumentText=stationRWYs;
        	
        	connOCC.Close();
			//Ignore box Ticked or not
			conn.Open();
			if(ChckBox_SeeIgnored.Checked)
			{
				query2 = "SELECT * FROM filteredNotams_table WHERE location=?";
			}
     		else
     		{
     			query2 = "SELECT * FROM filteredNotams_table WHERE (Status='K') AND (location=?)";
     		}
			OleDbCommand command4 = new OleDbCommand(query2, conn);
			command4.Parameters.AddWithValue("?", AP);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		int notam_ID=0;
     		string notam_key="";
     		string notam_text="";
     		string location="";
     		string fromDate="";
     		string fromDay="";
     		string fromMonth="";
     		string fromYear="";
     		string fromHour="";
     		string fromDateText="";
     		string tillDate="";
     		string tillDay="";
     		string tillMonth="";
     		string tillYear="";
     		string tillHour="";
     		string tillDateText="";
     		int nbNotams = 0;
     		int Top = 0;
     		Button[] keep_Buttons = new Button[20000];
     		//Label[] lbl_notam_text= new Label[20000];
     		RichTextBox[] RchTxt_notam_text = new RichTextBox[20000];
			CheckBox[] apt_CLSD_Chckbox= new CheckBox[20000];
			CheckBox[] apt_CATI_Chckbox= new CheckBox[20000];
			CheckBox[] apt_NILS_Chckbox= new CheckBox[20000];
			CheckBox[] apt_NOALTN_Chckbox= new CheckBox[20000];
			CheckBox[] apt_FUEL_Chckbox= new CheckBox[20000];
			CheckBox[] apt_MISC_Chckbox= new CheckBox[20000];
			CheckBox[] apt_AIPSUP_Chckbox= new CheckBox[20000];
			CheckBox[] apt_RWYCLSD_Chckbox= new CheckBox[20000];
			TextBox[] remark_Txtbox= new TextBox[20000];
			Button[] remark_Buttons = new Button[20000];
     		string Status="";
     		string Impact="";
     		string Remark="";
     		Top=100;
        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(0)) notam_ID=dBreader.GetInt32(0);
            		if(!dBreader.IsDBNull(5)) fromDate=dBreader.GetString(5);
            		if(!dBreader.IsDBNull(6)) tillDate=dBreader.GetString(6);
            		if(!dBreader.IsDBNull(7)) notam_text=dBreader.GetString(7);
            		if(!dBreader.IsDBNull(8)) location=dBreader.GetString(8);
            		if(!dBreader.IsDBNull(10)) notam_key=dBreader.GetString(10);
            		if(!dBreader.IsDBNull(12)) Status=dBreader.GetString(12);
            		if(!dBreader.IsDBNull(13)) Impact=dBreader.GetString(13);
            		if(!dBreader.IsDBNull(14)) Remark=dBreader.GetString(14);
            		notam_text=notam_text.Replace("(char)39","'");
            		
            		Label lbl_notam_key = new Label();
            		FontFamily courier =  new FontFamily("Courier New");
					lbl_notam_key.Font = new Font(courier, 10, FontStyle.Regular);
					lbl_notam_key.Tag ="dispose";
					lbl_notam_key.Top = Top;
     				lbl_notam_key.Size = new Size(125, 16);
     				lbl_notam_key.ForeColor = Color.Black;
     				lbl_notam_key.Text = notam_key;
     				lbl_notam_key.Left = 210;
     				tabPage2.Controls.Add(lbl_notam_key);
            		
     				fromDate=fromDate.Substring(0,16);
     				fromHour=fromDate.Substring(11,5);
     				fromDay=fromDate.Substring(8,2);
     				fromMonth=fromDate.Substring(5,2);
     				if(fromMonth=="01")fromMonth="JAN";
     				if(fromMonth=="02")fromMonth="FEB";
     				if(fromMonth=="03")fromMonth="MAR";
     				if(fromMonth=="04")fromMonth="APR";
     				if(fromMonth=="05")fromMonth="MAY";
     				if(fromMonth=="06")fromMonth="JUN";
     				if(fromMonth=="07")fromMonth="JUL";
     				if(fromMonth=="08")fromMonth="AUG";
     				if(fromMonth=="09")fromMonth="SEP";
     				if(fromMonth=="10")fromMonth="OCT";
     				if(fromMonth=="11")fromMonth="NOV";
     				if(fromMonth=="12")fromMonth="DEC";
     				fromYear=fromDate.Substring(0,4);
     				fromDateText=fromDay+fromMonth+fromYear+"("+fromHour+")";
     				Label lbl_notam_from = new Label();
            		//FontFamily courier =  new FontFamily("Courier New");
					lbl_notam_from.Font = new Font(courier, 10, FontStyle.Regular);
					lbl_notam_from.Tag ="dispose";
					lbl_notam_from.Top = Top;
     				lbl_notam_from.Size = new Size(200, 16);
     				lbl_notam_from.ForeColor = Color.Black;
     				lbl_notam_from.Text = "From : "+ fromDateText;
     				lbl_notam_from.Left = 335;
     				tabPage2.Controls.Add(lbl_notam_from);
     				
     				tillDate=tillDate.Substring(0,16);
     				tillHour=tillDate.Substring(11,5);
     				tillDay=tillDate.Substring(8,2);
     				tillMonth=tillDate.Substring(5,2);
     				if(tillMonth=="01")tillMonth="JAN";
     				if(tillMonth=="02")tillMonth="FEB";
     				if(tillMonth=="03")tillMonth="MAR";
     				if(tillMonth=="04")tillMonth="APR";
     				if(tillMonth=="05")tillMonth="MAY";
     				if(tillMonth=="06")tillMonth="JUN";
     				if(tillMonth=="07")tillMonth="JUL";
     				if(tillMonth=="08")tillMonth="AUG";
     				if(tillMonth=="09")tillMonth="SEP";
     				if(tillMonth=="10")tillMonth="OCT";
     				if(tillMonth=="11")tillMonth="NOV";
     				if(tillMonth=="12")tillMonth="DEC";
     				tillYear=tillDate.Substring(0,4);
     				tillDateText=tillDay+tillMonth+tillYear+"("+tillHour+")";
     				Label lbl_notam_till = new Label();
            		//FontFamily courier =  new FontFamily("Courier New");
					lbl_notam_till.Font = new Font(courier, 10, FontStyle.Regular);
					lbl_notam_till.Tag ="dispose";
					lbl_notam_till.Top = Top;
     				lbl_notam_till.Size = new Size(200, 16);
     				lbl_notam_till.ForeColor = Color.Black;
     				lbl_notam_till.Text = "Till : "+ tillDateText;
     				lbl_notam_till.Left = 540;
     				tabPage2.Controls.Add(lbl_notam_till);
     				
//            		int height = notam_text.Length/50*20;
//            		if(height<80)height=80;
//            		lbl_notam_text[notam_ID] = new Label();
//            		//FontFamily courier =  new FontFamily("Courier New");
//					lbl_notam_text[notam_ID].Font = new Font(courier, 10, FontStyle.Regular);
//					lbl_notam_text[notam_ID].Tag ="dispose";
//					lbl_notam_text[notam_ID].Top = Top+20;
//     				lbl_notam_text[notam_ID].Size = new Size(550, height);
//     				lbl_notam_text[notam_ID].ForeColor = Color.Black;
//     				lbl_notam_text[notam_ID].BackColor= Color.LightCoral;
//     				if(Status=="K")lbl_notam_text[notam_ID].BackColor= Color.CornflowerBlue;
//     				lbl_notam_text[notam_ID].Text = notam_text;
//     				lbl_notam_text[notam_ID].Left = 210;
//     				tabPage2.Controls.Add(lbl_notam_text[notam_ID] );
     						
     				int height = notam_text.Length/50*20;
            		if(height<80)height=80;
            		RchTxt_notam_text[notam_ID] = new RichTextBox();
            		//FontFamily courier =  new FontFamily("Courier New");
					RchTxt_notam_text[notam_ID].Font = new Font(courier, 10, FontStyle.Regular);
					RchTxt_notam_text[notam_ID].Tag ="dispose";
					RchTxt_notam_text[notam_ID].Top = Top+20;
     				RchTxt_notam_text[notam_ID].Size = new Size(550, height);
     				RchTxt_notam_text[notam_ID].ForeColor = Color.Black;
     				RchTxt_notam_text[notam_ID].BackColor= Color.LightCoral;
     				if(Status=="K")RchTxt_notam_text[notam_ID].BackColor= Color.CornflowerBlue;
     				RchTxt_notam_text[notam_ID].Text = notam_text;
     				RchTxt_notam_text[notam_ID].Left = 210;
     				RchTxt_notam_text[notam_ID].ReadOnly=true;
     				tabPage2.Controls.Add(RchTxt_notam_text[notam_ID] );		
     						
     				keep_Buttons[notam_ID] = new Button();
     				keep_Buttons[notam_ID].Tag = "dispose";
     				keep_Buttons[notam_ID].Size = new Size(40 , 25);
     				if(Status=="K")keep_Buttons[notam_ID].Size = new Size(50 , 25);
        			keep_Buttons[notam_ID].Location = new Point(770, Top+20);
        			int newSize =7;
        			keep_Buttons[notam_ID].Text = "Keep";
        			if(Status=="K")keep_Buttons[notam_ID].Text = "Ignore";
        			keep_Buttons[notam_ID].BackColor = Color.CornflowerBlue;
        			if(Status=="K")keep_Buttons[notam_ID].BackColor = Color.LightCoral;
        			keep_Buttons[notam_ID].Font = new Font(keep_Buttons[notam_ID].Font.FontFamily, newSize);
        			int i_notamToKeep = notam_ID;
        			if(Status=="")keep_Buttons[notam_ID].Click += (sender1, ex) => this.Keep_Notam(i_notamToKeep);
        			if(Status=="K")keep_Buttons[notam_ID].Click += (sender1, ex) => this.Ignore_Notam(i_notamToKeep);
     				tabPage2.Controls.Add(keep_Buttons[notam_ID]);	
     				
     				if(Status=="K"){
     				apt_CLSD_Chckbox[notam_ID] = new CheckBox();
     				apt_CLSD_Chckbox[notam_ID].Tag ="dispose";
     				apt_CLSD_Chckbox[notam_ID].Top = Top+44;
     				apt_CLSD_Chckbox[notam_ID].Left = 770;
     				apt_CLSD_Chckbox[notam_ID].Text = "APT CLSD";
     				apt_CLSD_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamAPTCLSD=notam_ID;
     				if(Impact=="A")apt_CLSD_Chckbox[notam_ID].Checked=true;
					else apt_CLSD_Chckbox[notam_ID].Checked=false;
					string A = "A";
     				apt_CLSD_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTCLSD, A);
     				tabPage2.Controls.Add(apt_CLSD_Chckbox[notam_ID]);
     				
     				apt_CATI_Chckbox[notam_ID] = new CheckBox();
     				apt_CATI_Chckbox[notam_ID].Tag ="dispose";
     				apt_CATI_Chckbox[notam_ID].Top = Top+44;
     				apt_CATI_Chckbox[notam_ID].Left = 850;
     				apt_CATI_Chckbox[notam_ID].Text = "APT CATI";
     				apt_CATI_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamAPTCATI=notam_ID;
     				if(Impact=="C")apt_CATI_Chckbox[notam_ID].Checked=true;
					else apt_CATI_Chckbox[notam_ID].Checked=false;
					string C ="C";
     				apt_CATI_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTCATI, C);
     				tabPage2.Controls.Add(apt_CATI_Chckbox[notam_ID]);
     				
     				apt_NILS_Chckbox[notam_ID] = new CheckBox();
     				apt_NILS_Chckbox[notam_ID].Tag ="dispose";
     				apt_NILS_Chckbox[notam_ID].Top = Top+44;
     				apt_NILS_Chckbox[notam_ID].Left = 940;
     				apt_NILS_Chckbox[notam_ID].Text = "No ILS";
     				apt_NILS_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamAPTNILS=notam_ID;
     				if(Impact=="N")apt_NILS_Chckbox[notam_ID].Checked=true;
					else apt_NILS_Chckbox[notam_ID].Checked=false;
					string N="N";
     				apt_NILS_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTNILS, N);
     				tabPage2.Controls.Add(apt_NILS_Chckbox[notam_ID]);
     				
     				apt_NOALTN_Chckbox[notam_ID] = new CheckBox();
     				apt_NOALTN_Chckbox[notam_ID].Tag ="dispose";
     				apt_NOALTN_Chckbox[notam_ID].Top = Top+68;
     				apt_NOALTN_Chckbox[notam_ID].Left = 770;
     				apt_NOALTN_Chckbox[notam_ID].Text = "Not ALTN";
     				apt_NOALTN_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamAPTNOALTN=notam_ID;
     				if(Impact=="D")apt_NOALTN_Chckbox[notam_ID].Checked=true;
					else apt_NOALTN_Chckbox[notam_ID].Checked=false;
					string D="D";
     				apt_NOALTN_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTNOALTN, D);
     				tabPage2.Controls.Add(apt_NOALTN_Chckbox[notam_ID]);
     				
     				apt_FUEL_Chckbox[notam_ID] = new CheckBox();
     				apt_FUEL_Chckbox[notam_ID].Tag ="dispose";
     				apt_FUEL_Chckbox[notam_ID].Top = Top+68;
     				apt_FUEL_Chckbox[notam_ID].Left = 850;
     				apt_FUEL_Chckbox[notam_ID].Text = "Fuel";
     				apt_FUEL_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamAPTFUEL=notam_ID;
     				if(Impact=="F")apt_FUEL_Chckbox[notam_ID].Checked=true;
					else apt_FUEL_Chckbox[notam_ID].Checked=false;
					string F="F";
     				apt_FUEL_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAPTFUEL, F);
     				tabPage2.Controls.Add(apt_FUEL_Chckbox[notam_ID]);
     				
     				apt_MISC_Chckbox[notam_ID] = new CheckBox();
     				apt_MISC_Chckbox[notam_ID].Tag ="dispose";
     				apt_MISC_Chckbox[notam_ID].Top = Top+68;
     				apt_MISC_Chckbox[notam_ID].Left = 940;
     				apt_MISC_Chckbox[notam_ID].Text = "MISC";
     				apt_MISC_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamMISC=notam_ID;
     				if(Impact=="M")apt_MISC_Chckbox[notam_ID].Checked=true;
					else apt_MISC_Chckbox[notam_ID].Checked=false;
					string M="M";
     				apt_MISC_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamMISC, M);
     				tabPage2.Controls.Add(apt_MISC_Chckbox[notam_ID]);
     				
     				//*************************************
     				apt_AIPSUP_Chckbox[notam_ID] = new CheckBox();
     				apt_AIPSUP_Chckbox[notam_ID].Tag ="dispose";
     				apt_AIPSUP_Chckbox[notam_ID].Top = Top+44;
     				apt_AIPSUP_Chckbox[notam_ID].Left = 1030;
     				apt_AIPSUP_Chckbox[notam_ID].Text = "SUP";
     				apt_AIPSUP_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamAIPSUP=notam_ID;
     				if(Impact=="AS")apt_AIPSUP_Chckbox[notam_ID].Checked=true;
					else apt_AIPSUP_Chckbox[notam_ID].Checked=false;
					string AS="AS";
     				apt_AIPSUP_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamAIPSUP, AS);
     				tabPage2.Controls.Add(apt_AIPSUP_Chckbox[notam_ID]);
     				
     				apt_RWYCLSD_Chckbox[notam_ID] = new CheckBox();
     				apt_RWYCLSD_Chckbox[notam_ID].Tag ="dispose";
     				apt_RWYCLSD_Chckbox[notam_ID].Top = Top+68;
     				apt_RWYCLSD_Chckbox[notam_ID].Left = 1030;
     				apt_RWYCLSD_Chckbox[notam_ID].Text = "RWY";
     				apt_RWYCLSD_Chckbox[notam_ID].Size = new Size (80,25);
     				int i_notamRWYCLSD=notam_ID;
     				if(Impact=="R")apt_RWYCLSD_Chckbox[notam_ID].Checked=true;
					else apt_RWYCLSD_Chckbox[notam_ID].Checked=false;
					string R="R";
     				apt_RWYCLSD_Chckbox[notam_ID].CheckedChanged += (sender1, ex) => this.Impact_Notam(i_notamRWYCLSD, R);
     				tabPage2.Controls.Add(apt_RWYCLSD_Chckbox[notam_ID]);
     				
     				
     				if(Impact=="A" || Impact=="C" || Impact=="N"|| Impact=="D"|| Impact=="F"|| Impact=="M"|| Impact=="AS"|| Impact=="R")
     				{
     					remark_Txtbox[notam_ID] = new TextBox();
     					remark_Txtbox[notam_ID].Tag="dispose";
     					remark_Txtbox[notam_ID].Top = Top+94;
     					remark_Txtbox[notam_ID].Left = 770;
     					remark_Txtbox[notam_ID].Size = new Size(250, 24);
     					remark_Txtbox[notam_ID].Text= Remark;
     					
     					remark_Buttons[notam_ID] = new Button();
     					remark_Buttons[notam_ID].Tag="dispose";
     					remark_Buttons[notam_ID].Top = Top+92;
     					remark_Buttons[notam_ID].Left = 1020;
     					remark_Buttons[notam_ID].Size = new Size(40, 24);
     					remark_Buttons[notam_ID].Text= "OK";
     					int i_remark=notam_ID;
     					
     					//remark_Txtbox[notam_ID].Leave += (sender1, ex) => this.Remark_Notam(i_remark, remark_Txtbox[i_remark].Text);
     					remark_Buttons[notam_ID].Click += (sender1, ex) => this.Remark_Notam(i_remark, remark_Txtbox[i_remark].Text);
     					tabPage2.Controls.Add(remark_Buttons[notam_ID]);
     					tabPage2.Controls.Add(remark_Txtbox[notam_ID]);
     				}
     				}
     				Top = Top+height+30;
     				nbNotams++;
     				Status="";
     				Impact="";
     				Remark="";
     				//A="";C="";N="";
            	}       		
        	}
        	conn.Close();
		}
		void Impact_Notam(int notam_ID, string I)
		{
			string stringIDNotam =notam_ID.ToString();
			string Imp = I;
			//RchTxtCSV.Text=stringIDNotam +"Letter : "+ Imp;
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
			conn.Open();
			string Impact="";
			var query2= "SELECT * FROM filteredNotams_table WHERE ID=?";

			OleDbCommand command4 = new OleDbCommand(query2, conn);
			command4.Parameters.AddWithValue("?", notam_ID);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		
        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(13)) Impact=dBreader.GetString(13);
            	}
        	}
        	conn.Close();
        	
        	//RchTxtCSV.Text=stringIDNotam+" "+Impact;
			conn.Open();
			var update ="";
			if(Impact=="")
				update = "UPDATE filteredNotams_table SET Impact=? WHERE ID=?";
			else
				update = "UPDATE filteredNotams_table SET Impact='' WHERE ID=?";

     		OleDbCommand command = new OleDbCommand(update, conn);
     		if(Impact=="") command.Parameters.AddWithValue("?", I);
     		command.Parameters.AddWithValue("?", notam_ID);
     		command.ExecuteNonQuery();
     		conn.Close();
     		ICAO_Notams();
     		Filter_Notams();
		}
		
		void Btn_ICAOClick(object sender, EventArgs e)
		{
			ICAO_Notams();
		}
		void ChckBox_SeeIgnoredCheckedChanged(object sender, EventArgs e)
		{
			ICAO_Notams();
		}
		void Remark_Notam(int notam_ID, string remark)
		{
			//RchTxtCSV.Text= notam_ID +" : "+remark;
			string stringIDNotam =notam_ID.ToString();
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= ICAO_storedNotams.mdb";
			conn.Open();
			
			var update = "UPDATE filteredNotams_table SET Remark=? WHERE ID=?";

     		OleDbCommand command = new OleDbCommand(update, conn);
     		command.Parameters.AddWithValue("?", remark);
     		command.Parameters.AddWithValue("?", notam_ID);
     		command.ExecuteNonQuery();
     		conn.Close();
     		ICAO_Notams();
			
			
		}
//		void Btn_updateDBClick(object sender, EventArgs e)
//		{
//			
//			// Downloading XML Notams from Web Service
//    		MessageBox.Show("Downloading XML Notams from Web Service...", "Info");
//    		GetXML();
//
//    		// Deleting withdrawned Notams from ICAO_storedNOTAMS.mdb
//    		MessageBox.Show("Deleting withdrawn NOTAMs from ICAO_storedNOTAMS.mdb...", "Info");
//    		deleteWithdrawnedNotams();
//
//		    // Adding New Notams to the Filtered NOTAMS table
//		    MessageBox.Show("Adding new NOTAMs to the Filtered NOTAMS table...", "Info");
//		    NewNotams();
//		
//		    // Deleting Notams from the Filtered NOTAMS table
//		    MessageBox.Show("Deleting old NOTAMs from the Filtered NOTAMS table...", "Info");
//		    DelOld();
//		
//		    // DB Updated
//		    MessageBox.Show("Database successfully updated!", "Done");
//
//		}

		
		public void ShowAutoPopup(string message, int durationMs = 1200)
		{
		    Form popup = new Form();
		    popup.StartPosition = FormStartPosition.CenterScreen;
		    popup.FormBorderStyle = FormBorderStyle.FixedToolWindow;
		    popup.Width = 350;
		    popup.Height = 120;
		    popup.TopMost = true;
		    popup.ControlBox = false;
		
		    Label lbl = new Label();
		    lbl.Dock = DockStyle.Fill;
		    lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		    lbl.Text = message;
		    lbl.Font = new System.Drawing.Font("Segoe UI", 10);
		    popup.Controls.Add(lbl);
		
		    popup.Shown += (s, e) =>
		    {
		        var timer = new System.Windows.Forms.Timer();
		        timer.Interval = durationMs;
		        timer.Tick += (s2, e2) =>
		        {
		            timer.Stop();
		            popup.Close();
		        };
		        timer.Start();
		    };
		
		    popup.Show();
		    popup.Refresh();
		}

		
		void Btn_updateDBClick(object sender, EventArgs e)
		{
		    // Downloading XML Notams from Web Service
		    ShowAutoPopup("Downloading XML Notams from Web Service...");
		    GetXML();
		
		    // Deleting withdrawn Notams
		    ShowAutoPopup("Deleting withdrawn NOTAMs...");
		    deleteWithdrawnedNotams();
		
		    // Adding new Notams
		    ShowAutoPopup("Adding new NOTAMs...");
		    NewNotams();
		
		    // Deleting old Notams
		    ShowAutoPopup("Deleting old NOTAMs...");
		    DelOld();
		
		    // DB Updated (popup finale cliquable)
		    MessageBox.Show("Database successfully updated!", "DB Updated");
		}

		
		void tab_RWYs()
		{

			tabPage5.VerticalScroll.Value = 0;
			
			List<Label> itemsToRemove = new List<Label>();
			foreach (Label label in tabPage5.Controls.OfType<Label>())
			{
    			if (label.Tag != null && label.Tag.ToString() == "dispose")
    			{
      		  		itemsToRemove.Add(label);
    			}
			}
			foreach (Label label in itemsToRemove)
			{
    			tabPage5.Controls.Remove(label);
    			label.Dispose();
			}
			
			List<TextBox> txtboxToRemove = new List<TextBox>();
			foreach (TextBox txtbox in tabPage5.Controls.OfType<TextBox>())
			{
    			if (txtbox.Tag != null && txtbox.Tag.ToString() == "dispose")
    			{
      		  		txtboxToRemove.Add(txtbox);
    			}
			}
			foreach (TextBox txtbox in txtboxToRemove)
			{
    			tabPage5.Controls.Remove(txtbox);
    			txtbox.Dispose();
			}
			
			List<RichTextBox> rchtxtboxToRemove = new List<RichTextBox>();
			foreach (RichTextBox rchtxtbox in tabPage5.Controls.OfType<RichTextBox>())
			{
    			if (rchtxtbox.Tag != null && rchtxtbox.Tag.ToString() == "dispose")
    			{
      		  		rchtxtboxToRemove.Add(rchtxtbox);
    			}
			}
			foreach (RichTextBox rchtxtbox in rchtxtboxToRemove)
			{
    			tabPage5.Controls.Remove(rchtxtbox);
    			rchtxtbox.Dispose();
			}
			
			List<CheckBox> chckboxToRemove = new List<CheckBox>();
			foreach (CheckBox chckbox in tabPage5.Controls.OfType<CheckBox>())
			{
    			if (chckbox.Tag != null && chckbox.Tag.ToString() == "dispose")
    			{
      		  		chckboxToRemove.Add(chckbox);
    			}
			}
			foreach (CheckBox chckbox in chckboxToRemove)
			{
    			tabPage5.Controls.Remove(chckbox);
    			chckbox.Dispose();
			}
			
			List<Button> buttonsToRemove = new List<Button>();
			foreach (Button button in tabPage5.Controls.OfType<Button>())
			{
    			if (button.Tag != null && button.Tag.ToString() == "dispose")
    			{
      		  		buttonsToRemove.Add(button);
    			}
			}
			foreach (Button button in buttonsToRemove)
			{
    			tabPage5.Controls.Remove(button);
    			button.Dispose();
			}
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			
			string AP="";
			
						
			conn.Open();
			var query2 = "SELECT * FROM Stations_ICAO_IATA ORDER BY ICAO";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		OleDbDataReader dBreader = command4.ExecuteReader();
     		int AP_ID=0;
     		string ICAO="";
     		string RWYs="";
     		string testICAO="";
     		int Top = 0;
     					
     		Button[] update_Buttons = new Button[20000];
     		RichTextBox[] RchTxt_RWYs= new RichTextBox[20000];
			//Label[] ICAO= new Label[20000];
     		
     		
     		Top=100;
        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(0)) AP_ID=dBreader.GetInt32(0);
            		if(!dBreader.IsDBNull(1)) ICAO=dBreader.GetString(1);
            		if(!dBreader.IsDBNull(6)) RWYs=dBreader.GetString(6);
            		testICAO+=ICAO+", ";
            		Label lbl_ICAO = new Label();
            		FontFamily courier =  new FontFamily("Courier New");
					lbl_ICAO.Font = new Font(courier, 10, FontStyle.Regular);
					lbl_ICAO.Tag ="dispose";
					lbl_ICAO.Top = Top;
     				lbl_ICAO.Size = new Size(125, 16);
     				lbl_ICAO.ForeColor = Color.Black;
     				lbl_ICAO.Text = ICAO;
     				lbl_ICAO.Left = 210;
     				tabPage5.Controls.Add(lbl_ICAO);
     				
     				int height = 100;
            		RchTxt_RWYs[AP_ID] = new RichTextBox();
            		//FontFamily courier =  new FontFamily("Courier New");
					RchTxt_RWYs[AP_ID].Font = new Font(courier, 10, FontStyle.Regular);
					RchTxt_RWYs[AP_ID].Tag ="dispose";
					RchTxt_RWYs[AP_ID].Top = Top+20;
     				RchTxt_RWYs[AP_ID].Size = new Size(550, height);
     				RchTxt_RWYs[AP_ID].ForeColor = Color.Black;
     				RchTxt_RWYs[AP_ID].BackColor= Color.White;
     				RchTxt_RWYs[AP_ID].Text = RWYs;
     				RchTxt_RWYs[AP_ID].ReadOnly=true;
     				RchTxt_RWYs[AP_ID].Left = 210;
    
     				tabPage5.Controls.Add(RchTxt_RWYs[AP_ID]);
     						
     				update_Buttons[AP_ID] = new Button();
     				update_Buttons[AP_ID].Tag = "dispose";
     				update_Buttons[AP_ID].Size = new Size(40 , 25);
        			update_Buttons[AP_ID].Location = new Point(770, Top+20);
        			int newSize =7;
        			update_Buttons[AP_ID].Text = "Edit";
        			update_Buttons[AP_ID].Font = new Font(update_Buttons[AP_ID].Font.FontFamily, newSize);
        			int i_AP= AP_ID;
        			//string RWYS_text=RchTxt_RWYs[i_AP].Text;
        			update_Buttons[AP_ID].Click += (sender1, ex) => this.Update_RWYs(i_AP);

     				tabPage5.Controls.Add(update_Buttons[AP_ID]);	
     				
     				Top = Top+height+30;
     				RWYs="";
            	}       		
        	}
        	conn.Close();
        	//RchTxtCSV.Text=testICAO;
		}
		void Update_RWYs(int AP_ID)
		{
			string RWYs="";
			string ICAO="";
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			

			conn.Open();
			var query2 = "SELECT * FROM Stations_ICAO_IATA WHERE ID=?";
     		OleDbCommand command4 = new OleDbCommand(query2, conn);
     		command4.Parameters.AddWithValue("?", AP_ID);
     		OleDbDataReader dBreader = command4.ExecuteReader();

        	if (dBreader.HasRows)
        	{
            	while (dBreader.Read())
            	{
            		if(!dBreader.IsDBNull(1)) ICAO=dBreader.GetString(1);
            		if(!dBreader.IsDBNull(6)) RWYs=dBreader.GetString(6);
            	}
        	}
			Lbl_ICAO_RWYs.Text = ICAO;
			RchTxt_updateRWYs.Text = RWYs;

     		conn.Close();
     		tab_RWYs();	
		}
		void Btn_updateRWysClick(object sender, EventArgs e)
		{
			string ICAO = Lbl_ICAO_RWYs.Text;
			string RWYs = RchTxt_updateRWYs.Text;
			
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			
			var updatelog = "UPDATE Stations_ICAO_IATA SET RWYs=? WHERE ICAO=?";

			conn.Open();
			OleDbCommand commandeinsert = new OleDbCommand(updatelog, conn);
			commandeinsert.Parameters.AddWithValue("?", RWYs);
			commandeinsert.Parameters.AddWithValue("?", ICAO);
			// Execution
			commandeinsert.ExecuteNonQuery();
			conn.Close();
			tab_RWYs();
					
		}
		void Btn_reportClick(object sender, EventArgs e)
		{
			Report();
		}
		void Btn_delWithdrawnedClick(object sender, EventArgs e)
		{
			deleteWithdrawnedNotams();
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
			
			//var opsDate = DateTime.Now;
			//string stringFrom = opsDate.ToShortDateString();
			//string stringClosed = "No";
			// Connexion à la DB message.mdb
			//try{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.JET.OLEDB.4.0;" + @"Data source= OCC.mdb";
			conn.Open();
			if(Btn_addAPT.Text=="Edit")
			{
				Btn_addAPT.Text="Add Airport !";
				string editID=Btn_addAPT.Tag.ToString();
				int intID=int.Parse(editID);
				var updatelog = "UPDATE Stations_ICAO_IATA SET ICAO=?,IATA=?,LH=?,FedEx=?,Charters=? WHERE ID=?";

				OleDbCommand commandeinsert = new OleDbCommand(updatelog, conn);
				commandeinsert.Parameters.AddWithValue("?", stringICAO);
				commandeinsert.Parameters.AddWithValue("?", stringIATA);
				commandeinsert.Parameters.AddWithValue("?", stringLongHaul);
				commandeinsert.Parameters.AddWithValue("?", stringFedEx);
				commandeinsert.Parameters.AddWithValue("?", stringCharters);
				commandeinsert.Parameters.AddWithValue("?", intID);
				// Execution
				commandeinsert.ExecuteNonQuery();
				Airport_List();
			}
			else
			{
				var insertlog = "INSERT INTO Stations_ICAO_IATA ([ICAO], [IATA], [LH], [FedEx], [Charters]) VALUES (?,?,?,?,?)";

				OleDbCommand commandeinsert = new OleDbCommand(insertlog, conn);
				commandeinsert.Parameters.AddWithValue("?", stringICAO);
				commandeinsert.Parameters.AddWithValue("?", stringIATA);
				commandeinsert.Parameters.AddWithValue("?", stringLongHaul);
				commandeinsert.Parameters.AddWithValue("?", stringFedEx);
				commandeinsert.Parameters.AddWithValue("?", stringCharters);
				// Execution
				commandeinsert.ExecuteNonQuery();
				Airport_List();
			}
				
				conn.Close();
				LoadStationsCache();
				//}
				//catch(Exception Ex)
				//{
					//MessageBox.Show("Could not update the database. Your record has not been saved. If error persist, contact the administrator.", "Access database issue", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				//}


				Airport_List();
		}
		void Btn_printReportClick(object sender, EventArgs e)
		{
			Web_report.Print();
		}
		void Btn_AIP_Sup_reportClick(object sender, EventArgs e)
		{
			Sup_Report();
		}
		void Btn_Sup_printReportClick(object sender, EventArgs e)
		{
			Web_Sup_report.Print();
		}
		void Btn_restartAppClick(object sender, EventArgs e)
		{
			Application.Restart();
		}
		void Btn_XMLClick(object sender, EventArgs e)
		{
			GetXML();
		}
		void Btn_reloadClick(object sender, EventArgs e)
		{
			Reload_text();
		}
		
	}
}
 