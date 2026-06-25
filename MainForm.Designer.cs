/*
 * Created by SharpDevelop.
 * User: jgazon
 * Date: 24-02-21
 * Time: 17:16
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace ICAO_CSV
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		private System.Windows.Forms.RichTextBox RchTxtCSV;
		private System.Windows.Forms.Button Btn_CSV;
		private System.Windows.Forms.ProgressBar progressBar;
		private System.Windows.Forms.Button Btn_split;
		private System.Windows.Forms.Button Btn_newNotams;
		private System.Windows.Forms.Button Btn_analyse;
		private System.Windows.Forms.Label Lbl_location;
		private System.Windows.Forms.Button Btn_delOld;
		private System.Windows.Forms.Label Lbl_notamsUnchecked;
		private System.Windows.Forms.Button Btn_analyzeNotams;
		private System.Windows.Forms.Button Btn_submitNotams;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.Button Btn_ICAO;
		private System.Windows.Forms.TextBox TxtBox_ICAO;
		private System.Windows.Forms.CheckBox ChckBox_SeeIgnored;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.Button Btn_updateDB;
		private System.Windows.Forms.TabPage tabPage4;
		private System.Windows.Forms.WebBrowser webBrowser2;
		private System.Windows.Forms.WebBrowser webBrowser1;
		private System.Windows.Forms.WebBrowser Web_report;
		private System.Windows.Forms.TabPage tabPage5;
		private System.Windows.Forms.RichTextBox RchTxt_updateRWYs;
		private System.Windows.Forms.Label Lbl_ICAO_RWYs;
		private System.Windows.Forms.Button Btn_updateRWys;
		private System.Windows.Forms.WebBrowser Web_FilterHeader;
		private System.Windows.Forms.Button btn_report;
		private System.Windows.Forms.RadioButton radBtn_24Hrs;
		private System.Windows.Forms.RadioButton radBtn_noFilter;
		private System.Windows.Forms.RadioButton radBtn_31days;
		private System.Windows.Forms.RadioButton radBtn_7days;
		private System.Windows.Forms.Button Btn_delWithdrawned;
		private System.Windows.Forms.TabPage APT_List;
		private System.Windows.Forms.Button Btn_CopyAPTList;
		private System.Windows.Forms.Button Btn_addAPT;
		private System.Windows.Forms.CheckBox ChckBx_APT_Charters;
		private System.Windows.Forms.CheckBox ChckBx_APT_FedEx;
		private System.Windows.Forms.CheckBox ChckBx_APT_LH;
		private System.Windows.Forms.TextBox TxtBox_APT_IATA;
		private System.Windows.Forms.TextBox TxtBox_APT_ICAO;
		private System.Windows.Forms.Button Btn_printReport;
		private System.Windows.Forms.Button Btn_exportReport;
		private System.Windows.Forms.TabPage AIP_SUP_report;
		private System.Windows.Forms.Button Btn_Sup_printReport;
		private System.Windows.Forms.Button Btn_Sup_exportReport;
		private System.Windows.Forms.RadioButton radBtn_Sup_31days;
		private System.Windows.Forms.RadioButton radBtn_Sup_7days;
		private System.Windows.Forms.RadioButton radBtn_Sup_24Hrs;
		private System.Windows.Forms.RadioButton radBtn_Sup_noFilter;
		private System.Windows.Forms.Button btn_AIP_Sup_report;
		private System.Windows.Forms.WebBrowser Web_Sup_report;
		private System.Windows.Forms.Button Btn_restartApp;
		private System.Windows.Forms.Button Btn_XML;
		private System.Windows.Forms.Button Btn_reload;
		private System.Windows.Forms.WebBrowser Web_ICAONotams;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.RchTxtCSV = new System.Windows.Forms.RichTextBox();
			this.Btn_CSV = new System.Windows.Forms.Button();
			this.progressBar = new System.Windows.Forms.ProgressBar();
			this.Btn_split = new System.Windows.Forms.Button();
			this.Btn_newNotams = new System.Windows.Forms.Button();
			this.Lbl_location = new System.Windows.Forms.Label();
			this.Btn_delOld = new System.Windows.Forms.Button();
			this.Lbl_notamsUnchecked = new System.Windows.Forms.Label();
			this.Btn_analyzeNotams = new System.Windows.Forms.Button();
			this.Btn_submitNotams = new System.Windows.Forms.Button();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.Web_FilterHeader = new System.Windows.Forms.WebBrowser();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.Web_ICAONotams = new System.Windows.Forms.WebBrowser();
			this.ChckBox_SeeIgnored = new System.Windows.Forms.CheckBox();
			this.Btn_ICAO = new System.Windows.Forms.Button();
			this.TxtBox_ICAO = new System.Windows.Forms.TextBox();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.Btn_reload = new System.Windows.Forms.Button();
			this.Btn_XML = new System.Windows.Forms.Button();
			this.Btn_restartApp = new System.Windows.Forms.Button();
			this.Btn_delWithdrawned = new System.Windows.Forms.Button();
			this.Btn_updateDB = new System.Windows.Forms.Button();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.Btn_printReport = new System.Windows.Forms.Button();
			this.Btn_exportReport = new System.Windows.Forms.Button();
			this.radBtn_31days = new System.Windows.Forms.RadioButton();
			this.radBtn_7days = new System.Windows.Forms.RadioButton();
			this.radBtn_24Hrs = new System.Windows.Forms.RadioButton();
			this.radBtn_noFilter = new System.Windows.Forms.RadioButton();
			this.btn_report = new System.Windows.Forms.Button();
			this.Web_report = new System.Windows.Forms.WebBrowser();
			this.webBrowser2 = new System.Windows.Forms.WebBrowser();
			this.webBrowser1 = new System.Windows.Forms.WebBrowser();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.RchTxt_updateRWYs = new System.Windows.Forms.RichTextBox();
			this.Lbl_ICAO_RWYs = new System.Windows.Forms.Label();
			this.Btn_updateRWys = new System.Windows.Forms.Button();
			this.APT_List = new System.Windows.Forms.TabPage();
			this.Btn_CopyAPTList = new System.Windows.Forms.Button();
			this.Btn_addAPT = new System.Windows.Forms.Button();
			this.ChckBx_APT_Charters = new System.Windows.Forms.CheckBox();
			this.ChckBx_APT_FedEx = new System.Windows.Forms.CheckBox();
			this.ChckBx_APT_LH = new System.Windows.Forms.CheckBox();
			this.TxtBox_APT_IATA = new System.Windows.Forms.TextBox();
			this.TxtBox_APT_ICAO = new System.Windows.Forms.TextBox();
			this.AIP_SUP_report = new System.Windows.Forms.TabPage();
			this.Btn_Sup_printReport = new System.Windows.Forms.Button();
			this.Btn_Sup_exportReport = new System.Windows.Forms.Button();
			this.radBtn_Sup_31days = new System.Windows.Forms.RadioButton();
			this.radBtn_Sup_7days = new System.Windows.Forms.RadioButton();
			this.radBtn_Sup_24Hrs = new System.Windows.Forms.RadioButton();
			this.radBtn_Sup_noFilter = new System.Windows.Forms.RadioButton();
			this.btn_AIP_Sup_report = new System.Windows.Forms.Button();
			this.Web_Sup_report = new System.Windows.Forms.WebBrowser();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.tabPage2.SuspendLayout();
			this.tabPage3.SuspendLayout();
			this.tabPage4.SuspendLayout();
			this.tabPage5.SuspendLayout();
			this.APT_List.SuspendLayout();
			this.AIP_SUP_report.SuspendLayout();
			this.SuspendLayout();
			// 
			// RchTxtCSV
			// 
			this.RchTxtCSV.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.RchTxtCSV.Location = new System.Drawing.Point(293, 17);
			this.RchTxtCSV.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.RchTxtCSV.Name = "RchTxtCSV";
			this.RchTxtCSV.Size = new System.Drawing.Size(1512, 793);
			this.RchTxtCSV.TabIndex = 0;
			this.RchTxtCSV.Text = "";
			// 
			// Btn_CSV
			// 
			this.Btn_CSV.Location = new System.Drawing.Point(17, 79);
			this.Btn_CSV.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_CSV.Name = "Btn_CSV";
			this.Btn_CSV.Size = new System.Drawing.Size(205, 27);
			this.Btn_CSV.TabIndex = 1;
			this.Btn_CSV.Text = "CSV";
			this.Btn_CSV.UseVisualStyleBackColor = true;
			this.Btn_CSV.Click += new System.EventHandler(this.Btn_CSVClick);
			// 
			// progressBar
			// 
			this.progressBar.Location = new System.Drawing.Point(17, 111);
			this.progressBar.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.progressBar.Name = "progressBar";
			this.progressBar.Size = new System.Drawing.Size(205, 14);
			this.progressBar.TabIndex = 2;
			// 
			// Btn_split
			// 
			this.Btn_split.Location = new System.Drawing.Point(17, 130);
			this.Btn_split.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_split.Name = "Btn_split";
			this.Btn_split.Size = new System.Drawing.Size(205, 27);
			this.Btn_split.TabIndex = 3;
			this.Btn_split.Text = "Split";
			this.Btn_split.UseVisualStyleBackColor = true;
			this.Btn_split.Click += new System.EventHandler(this.Btn_splitClick);
			// 
			// Btn_newNotams
			// 
			this.Btn_newNotams.Location = new System.Drawing.Point(17, 162);
			this.Btn_newNotams.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_newNotams.Name = "Btn_newNotams";
			this.Btn_newNotams.Size = new System.Drawing.Size(205, 27);
			this.Btn_newNotams.TabIndex = 4;
			this.Btn_newNotams.Text = "New NOTAMS";
			this.Btn_newNotams.UseVisualStyleBackColor = true;
			this.Btn_newNotams.Click += new System.EventHandler(this.Btn_newNotamsClick);
			// 
			// Lbl_location
			// 
			this.Lbl_location.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Lbl_location.Location = new System.Drawing.Point(680, 42);
			this.Lbl_location.Name = "Lbl_location";
			this.Lbl_location.Size = new System.Drawing.Size(128, 28);
			this.Lbl_location.TabIndex = 11;
			this.Lbl_location.Text = "Location";
			// 
			// Btn_delOld
			// 
			this.Btn_delOld.Location = new System.Drawing.Point(17, 194);
			this.Btn_delOld.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_delOld.Name = "Btn_delOld";
			this.Btn_delOld.Size = new System.Drawing.Size(205, 27);
			this.Btn_delOld.TabIndex = 14;
			this.Btn_delOld.Text = "Delete Old";
			this.Btn_delOld.UseVisualStyleBackColor = true;
			this.Btn_delOld.Click += new System.EventHandler(this.Btn_delOldClick);
			// 
			// Lbl_notamsUnchecked
			// 
			this.Lbl_notamsUnchecked.Location = new System.Drawing.Point(684, 78);
			this.Lbl_notamsUnchecked.Name = "Lbl_notamsUnchecked";
			this.Lbl_notamsUnchecked.Size = new System.Drawing.Size(260, 25);
			this.Lbl_notamsUnchecked.TabIndex = 15;
			this.Lbl_notamsUnchecked.Text = "Notams unchecked :";
			// 
			// Btn_analyzeNotams
			// 
			this.Btn_analyzeNotams.Location = new System.Drawing.Point(847, 6);
			this.Btn_analyzeNotams.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_analyzeNotams.Name = "Btn_analyzeNotams";
			this.Btn_analyzeNotams.Size = new System.Drawing.Size(221, 34);
			this.Btn_analyzeNotams.TabIndex = 16;
			this.Btn_analyzeNotams.Text = "Analyze New Notams !";
			this.Btn_analyzeNotams.UseVisualStyleBackColor = true;
			this.Btn_analyzeNotams.Click += new System.EventHandler(this.Btn_analyzeNotamsClick);
			// 
			// Btn_submitNotams
			// 
			this.Btn_submitNotams.Location = new System.Drawing.Point(811, 1159);
			this.Btn_submitNotams.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_submitNotams.Name = "Btn_submitNotams";
			this.Btn_submitNotams.Size = new System.Drawing.Size(133, 28);
			this.Btn_submitNotams.TabIndex = 19;
			this.Btn_submitNotams.Text = "Submit !";
			this.Btn_submitNotams.UseVisualStyleBackColor = true;
			this.Btn_submitNotams.Click += new System.EventHandler(this.Btn_submitNotamsClick);
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Controls.Add(this.tabPage3);
			this.tabControl1.Controls.Add(this.tabPage4);
			this.tabControl1.Controls.Add(this.tabPage5);
			this.tabControl1.Controls.Add(this.APT_List);
			this.tabControl1.Controls.Add(this.AIP_SUP_report);
			this.tabControl1.Location = new System.Drawing.Point(12, 10);
			this.tabControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(1885, 1268);
			this.tabControl1.TabIndex = 20;
			// 
			// tabPage1
			// 
			this.tabPage1.AutoScroll = true;
			this.tabPage1.Controls.Add(this.Web_FilterHeader);
			this.tabPage1.Controls.Add(this.Lbl_location);
			this.tabPage1.Controls.Add(this.Btn_submitNotams);
			this.tabPage1.Controls.Add(this.Btn_analyzeNotams);
			this.tabPage1.Controls.Add(this.Lbl_notamsUnchecked);
			this.tabPage1.Location = new System.Drawing.Point(4, 25);
			this.tabPage1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabPage1.Size = new System.Drawing.Size(1877, 1239);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Filter New Notams";
			this.tabPage1.UseVisualStyleBackColor = true;
			// 
			// Web_FilterHeader
			//
			this.Web_FilterHeader.Location = new System.Drawing.Point(7, 6);
			this.Web_FilterHeader.Name = "Web_FilterHeader";
			this.Web_FilterHeader.Size = new System.Drawing.Size(490, 90);
			this.Web_FilterHeader.TabIndex = 20;
			this.Web_FilterHeader.ScrollBarsEnabled = false;
			// 
			// tabPage2
			// 
			this.tabPage2.AutoScroll = true;
			this.tabPage2.Controls.Add(this.Web_ICAONotams);
			this.tabPage2.Controls.Add(this.ChckBox_SeeIgnored);
			this.tabPage2.Controls.Add(this.Btn_ICAO);
			this.tabPage2.Controls.Add(this.TxtBox_ICAO);
			this.tabPage2.Location = new System.Drawing.Point(4, 25);
			this.tabPage2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabPage2.Size = new System.Drawing.Size(1877, 1239);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Stations";
			this.tabPage2.UseVisualStyleBackColor = true;
			// 
			// Web_ICAONotams
			// 
			this.Web_ICAONotams.Location = new System.Drawing.Point(19, 34);
			this.Web_ICAONotams.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Web_ICAONotams.MinimumSize = new System.Drawing.Size(27, 25);
			this.Web_ICAONotams.Name = "Web_ICAONotams";
			this.Web_ICAONotams.Size = new System.Drawing.Size(233, 363);
			this.Web_ICAONotams.TabIndex = 4;
			// 
			// ChckBox_SeeIgnored
			// 
			this.ChckBox_SeeIgnored.Location = new System.Drawing.Point(573, 86);
			this.ChckBox_SeeIgnored.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.ChckBox_SeeIgnored.Name = "ChckBox_SeeIgnored";
			this.ChckBox_SeeIgnored.Size = new System.Drawing.Size(147, 27);
			this.ChckBox_SeeIgnored.TabIndex = 2;
			this.ChckBox_SeeIgnored.Text = "See \'Ignored\'";
			this.ChckBox_SeeIgnored.UseVisualStyleBackColor = true;
			this.ChckBox_SeeIgnored.CheckedChanged += new System.EventHandler(this.ChckBox_SeeIgnoredCheckedChanged);
			// 
			// Btn_ICAO
			// 
			this.Btn_ICAO.Location = new System.Drawing.Point(691, 34);
			this.Btn_ICAO.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_ICAO.Name = "Btn_ICAO";
			this.Btn_ICAO.Size = new System.Drawing.Size(60, 34);
			this.Btn_ICAO.TabIndex = 1;
			this.Btn_ICAO.Text = "OK";
			this.Btn_ICAO.UseVisualStyleBackColor = true;
			this.Btn_ICAO.Click += new System.EventHandler(this.Btn_ICAOClick);
			// 
			// TxtBox_ICAO
			// 
			this.TxtBox_ICAO.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.TxtBox_ICAO.Location = new System.Drawing.Point(569, 34);
			this.TxtBox_ICAO.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.TxtBox_ICAO.Name = "TxtBox_ICAO";
			this.TxtBox_ICAO.Size = new System.Drawing.Size(97, 38);
			this.TxtBox_ICAO.TabIndex = 0;
			this.TxtBox_ICAO.Text = "ICAO";
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.Btn_reload);
			this.tabPage3.Controls.Add(this.Btn_XML);
			this.tabPage3.Controls.Add(this.Btn_restartApp);
			this.tabPage3.Controls.Add(this.Btn_delWithdrawned);
			this.tabPage3.Controls.Add(this.Btn_updateDB);
			this.tabPage3.Controls.Add(this.RchTxtCSV);
			this.tabPage3.Controls.Add(this.Btn_CSV);
			this.tabPage3.Controls.Add(this.Btn_delOld);
			this.tabPage3.Controls.Add(this.progressBar);
			this.tabPage3.Controls.Add(this.Btn_newNotams);
			this.tabPage3.Controls.Add(this.Btn_split);
			this.tabPage3.Location = new System.Drawing.Point(4, 25);
			this.tabPage3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(1877, 1239);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "DB Update";
			this.tabPage3.UseVisualStyleBackColor = true;
			// 
			// Btn_reload
			// 
			this.Btn_reload.Location = new System.Drawing.Point(17, 550);
			this.Btn_reload.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_reload.Name = "Btn_reload";
			this.Btn_reload.Size = new System.Drawing.Size(205, 27);
			this.Btn_reload.TabIndex = 19;
			this.Btn_reload.Text = "Reload Text";
			this.Btn_reload.UseVisualStyleBackColor = true;
			this.Btn_reload.Click += new System.EventHandler(this.Btn_reloadClick);
			// 
			// Btn_XML
			// 
			this.Btn_XML.Location = new System.Drawing.Point(17, 444);
			this.Btn_XML.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_XML.Name = "Btn_XML";
			this.Btn_XML.Size = new System.Drawing.Size(205, 27);
			this.Btn_XML.TabIndex = 18;
			this.Btn_XML.Text = "XML";
			this.Btn_XML.UseVisualStyleBackColor = true;
			this.Btn_XML.Click += new System.EventHandler(this.Btn_XMLClick);
			// 
			// Btn_restartApp
			// 
			this.Btn_restartApp.Location = new System.Drawing.Point(17, 362);
			this.Btn_restartApp.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_restartApp.Name = "Btn_restartApp";
			this.Btn_restartApp.Size = new System.Drawing.Size(205, 27);
			this.Btn_restartApp.TabIndex = 17;
			this.Btn_restartApp.Text = "Restart App";
			this.Btn_restartApp.UseVisualStyleBackColor = true;
			this.Btn_restartApp.Click += new System.EventHandler(this.Btn_restartAppClick);
			// 
			// Btn_delWithdrawned
			// 
			this.Btn_delWithdrawned.Location = new System.Drawing.Point(17, 226);
			this.Btn_delWithdrawned.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_delWithdrawned.Name = "Btn_delWithdrawned";
			this.Btn_delWithdrawned.Size = new System.Drawing.Size(205, 27);
			this.Btn_delWithdrawned.TabIndex = 16;
			this.Btn_delWithdrawned.Text = "Delete Withdrawned";
			this.Btn_delWithdrawned.UseVisualStyleBackColor = true;
			this.Btn_delWithdrawned.Click += new System.EventHandler(this.Btn_delWithdrawnedClick);
			// 
			// Btn_updateDB
			// 
			this.Btn_updateDB.Location = new System.Drawing.Point(17, 17);
			this.Btn_updateDB.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_updateDB.Name = "Btn_updateDB";
			this.Btn_updateDB.Size = new System.Drawing.Size(205, 39);
			this.Btn_updateDB.TabIndex = 15;
			this.Btn_updateDB.Text = "Update DB";
			this.Btn_updateDB.UseVisualStyleBackColor = true;
			this.Btn_updateDB.Click += new System.EventHandler(this.Btn_updateDBClick);
			// 
			// tabPage4
			// 
			this.tabPage4.Controls.Add(this.Btn_printReport);
			this.tabPage4.Controls.Add(this.Btn_exportReport);
			this.tabPage4.Controls.Add(this.radBtn_31days);
			this.tabPage4.Controls.Add(this.radBtn_7days);
			this.tabPage4.Controls.Add(this.radBtn_24Hrs);
			this.tabPage4.Controls.Add(this.radBtn_noFilter);
			this.tabPage4.Controls.Add(this.btn_report);
			this.tabPage4.Controls.Add(this.Web_report);
			this.tabPage4.Controls.Add(this.webBrowser2);
			this.tabPage4.Controls.Add(this.webBrowser1);
			this.tabPage4.Location = new System.Drawing.Point(4, 25);
			this.tabPage4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabPage4.Name = "tabPage4";
			this.tabPage4.Size = new System.Drawing.Size(1877, 1239);
			this.tabPage4.TabIndex = 3;
			this.tabPage4.Text = "Report";
			this.tabPage4.UseVisualStyleBackColor = true;
			// 
			// Btn_printReport
			// 
			this.Btn_printReport.BackColor = System.Drawing.Color.Black;
			this.Btn_printReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Btn_printReport.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
			this.Btn_printReport.Location = new System.Drawing.Point(712, 5);
			this.Btn_printReport.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Btn_printReport.Name = "Btn_printReport";
			this.Btn_printReport.Size = new System.Drawing.Size(100, 28);
			this.Btn_printReport.TabIndex = 8;
			this.Btn_printReport.Text = "Print";
			this.Btn_printReport.UseVisualStyleBackColor = false;
			this.Btn_printReport.Click += new System.EventHandler(this.Btn_printReportClick);
			//
			// Btn_exportReport
			//
			this.Btn_exportReport.BackColor = System.Drawing.Color.DarkGreen;
			this.Btn_exportReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Btn_exportReport.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
			this.Btn_exportReport.Location = new System.Drawing.Point(822, 5);
			this.Btn_exportReport.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Btn_exportReport.Name = "Btn_exportReport";
			this.Btn_exportReport.Size = new System.Drawing.Size(100, 28);
			this.Btn_exportReport.TabIndex = 9;
			this.Btn_exportReport.Text = "Export PDF";
			this.Btn_exportReport.UseVisualStyleBackColor = false;
			this.Btn_exportReport.Click += new System.EventHandler(this.Btn_exportReportClick);
			//
			// radBtn_31days
			// 
			this.radBtn_31days.Location = new System.Drawing.Point(565, 4);
			this.radBtn_31days.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_31days.Name = "radBtn_31days";
			this.radBtn_31days.Size = new System.Drawing.Size(139, 30);
			this.radBtn_31days.TabIndex = 7;
			this.radBtn_31days.Text = "31 days";
			this.radBtn_31days.UseVisualStyleBackColor = true;
			// 
			// radBtn_7days
			// 
			this.radBtn_7days.Location = new System.Drawing.Point(447, 4);
			this.radBtn_7days.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_7days.Name = "radBtn_7days";
			this.radBtn_7days.Size = new System.Drawing.Size(139, 30);
			this.radBtn_7days.TabIndex = 6;
			this.radBtn_7days.Text = "7 days";
			this.radBtn_7days.UseVisualStyleBackColor = true;
			// 
			// radBtn_24Hrs
			// 
			this.radBtn_24Hrs.Location = new System.Drawing.Point(281, 4);
			this.radBtn_24Hrs.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_24Hrs.Name = "radBtn_24Hrs";
			this.radBtn_24Hrs.Size = new System.Drawing.Size(139, 30);
			this.radBtn_24Hrs.TabIndex = 5;
			this.radBtn_24Hrs.Text = "Next 24Hrs";
			this.radBtn_24Hrs.UseVisualStyleBackColor = true;
			// 
			// radBtn_noFilter
			// 
			this.radBtn_noFilter.Checked = true;
			this.radBtn_noFilter.Location = new System.Drawing.Point(149, 4);
			this.radBtn_noFilter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_noFilter.Name = "radBtn_noFilter";
			this.radBtn_noFilter.Size = new System.Drawing.Size(95, 30);
			this.radBtn_noFilter.TabIndex = 4;
			this.radBtn_noFilter.TabStop = true;
			this.radBtn_noFilter.Text = "No filter";
			this.radBtn_noFilter.UseVisualStyleBackColor = true;
			// 
			// btn_report
			// 
			this.btn_report.Location = new System.Drawing.Point(23, 4);
			this.btn_report.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.btn_report.Name = "btn_report";
			this.btn_report.Size = new System.Drawing.Size(100, 28);
			this.btn_report.TabIndex = 3;
			this.btn_report.Text = "Report !";
			this.btn_report.UseVisualStyleBackColor = true;
			this.btn_report.Click += new System.EventHandler(this.Btn_reportClick);
			// 
			// Web_report
			// 
			this.Web_report.Location = new System.Drawing.Point(0, 33);
			this.Web_report.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Web_report.MinimumSize = new System.Drawing.Size(20, 20);
			this.Web_report.Name = "Web_report";
			this.Web_report.Size = new System.Drawing.Size(1680, 1209);
			this.Web_report.TabIndex = 2;
			// 
			// webBrowser2
			// 
			this.webBrowser2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.webBrowser2.Location = new System.Drawing.Point(0, 0);
			this.webBrowser2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.webBrowser2.MinimumSize = new System.Drawing.Size(20, 20);
			this.webBrowser2.Name = "webBrowser2";
			this.webBrowser2.Size = new System.Drawing.Size(1877, 1239);
			this.webBrowser2.TabIndex = 1;
			// 
			// webBrowser1
			// 
			this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.webBrowser1.Location = new System.Drawing.Point(0, 0);
			this.webBrowser1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
			this.webBrowser1.Name = "webBrowser1";
			this.webBrowser1.Size = new System.Drawing.Size(1877, 1239);
			this.webBrowser1.TabIndex = 0;
			// 
			// tabPage5
			// 
			this.tabPage5.AutoScroll = true;
			this.tabPage5.Controls.Add(this.RchTxt_updateRWYs);
			this.tabPage5.Controls.Add(this.Lbl_ICAO_RWYs);
			this.tabPage5.Controls.Add(this.Btn_updateRWys);
			this.tabPage5.Location = new System.Drawing.Point(4, 25);
			this.tabPage5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Size = new System.Drawing.Size(1877, 1239);
			this.tabPage5.TabIndex = 4;
			this.tabPage5.Text = "RWYs";
			this.tabPage5.UseVisualStyleBackColor = true;
			// 
			// RchTxt_updateRWYs
			// 
			this.RchTxt_updateRWYs.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.RchTxt_updateRWYs.Location = new System.Drawing.Point(195, 23);
			this.RchTxt_updateRWYs.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.RchTxt_updateRWYs.Name = "RchTxt_updateRWYs";
			this.RchTxt_updateRWYs.Size = new System.Drawing.Size(813, 96);
			this.RchTxt_updateRWYs.TabIndex = 2;
			this.RchTxt_updateRWYs.Text = "";
			// 
			// Lbl_ICAO_RWYs
			// 
			this.Lbl_ICAO_RWYs.Location = new System.Drawing.Point(89, 26);
			this.Lbl_ICAO_RWYs.Name = "Lbl_ICAO_RWYs";
			this.Lbl_ICAO_RWYs.Size = new System.Drawing.Size(100, 23);
			this.Lbl_ICAO_RWYs.TabIndex = 1;
			this.Lbl_ICAO_RWYs.Text = "ICAO";
			// 
			// Btn_updateRWys
			// 
			this.Btn_updateRWys.Location = new System.Drawing.Point(1059, 23);
			this.Btn_updateRWys.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_updateRWys.Name = "Btn_updateRWys";
			this.Btn_updateRWys.Size = new System.Drawing.Size(75, 23);
			this.Btn_updateRWys.TabIndex = 0;
			this.Btn_updateRWys.Text = "OK";
			this.Btn_updateRWys.UseVisualStyleBackColor = true;
			this.Btn_updateRWys.Click += new System.EventHandler(this.Btn_updateRWysClick);
			// 
			// APT_List
			// 
			this.APT_List.Controls.Add(this.Btn_CopyAPTList);
			this.APT_List.Controls.Add(this.Btn_addAPT);
			this.APT_List.Controls.Add(this.ChckBx_APT_Charters);
			this.APT_List.Controls.Add(this.ChckBx_APT_FedEx);
			this.APT_List.Controls.Add(this.ChckBx_APT_LH);
			this.APT_List.Controls.Add(this.TxtBox_APT_IATA);
			this.APT_List.Controls.Add(this.TxtBox_APT_ICAO);
			this.APT_List.Location = new System.Drawing.Point(4, 25);
			this.APT_List.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.APT_List.Name = "APT_List";
			this.APT_List.Size = new System.Drawing.Size(1877, 1239);
			this.APT_List.TabIndex = 5;
			this.APT_List.Text = "APT List";
			this.APT_List.UseVisualStyleBackColor = true;
			// 
			// Btn_CopyAPTList
			// 
			this.Btn_CopyAPTList.Location = new System.Drawing.Point(907, 26);
			this.Btn_CopyAPTList.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_CopyAPTList.Name = "Btn_CopyAPTList";
			this.Btn_CopyAPTList.Size = new System.Drawing.Size(100, 27);
			this.Btn_CopyAPTList.TabIndex = 13;
			this.Btn_CopyAPTList.Text = "Copy List";
			this.Btn_CopyAPTList.UseVisualStyleBackColor = true;
			// 
			// Btn_addAPT
			// 
			this.Btn_addAPT.Location = new System.Drawing.Point(509, 26);
			this.Btn_addAPT.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Btn_addAPT.Name = "Btn_addAPT";
			this.Btn_addAPT.Size = new System.Drawing.Size(115, 27);
			this.Btn_addAPT.TabIndex = 12;
			this.Btn_addAPT.Text = "Add Airport !";
			this.Btn_addAPT.UseVisualStyleBackColor = true;
			this.Btn_addAPT.Click += new System.EventHandler(this.Btn_addAPTClick);
			// 
			// ChckBx_APT_Charters
			// 
			this.ChckBx_APT_Charters.Location = new System.Drawing.Point(391, 30);
			this.ChckBx_APT_Charters.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.ChckBx_APT_Charters.Name = "ChckBx_APT_Charters";
			this.ChckBx_APT_Charters.Size = new System.Drawing.Size(92, 20);
			this.ChckBx_APT_Charters.TabIndex = 11;
			this.ChckBx_APT_Charters.Text = "Charters";
			this.ChckBx_APT_Charters.UseVisualStyleBackColor = true;
			// 
			// ChckBx_APT_FedEx
			// 
			this.ChckBx_APT_FedEx.Location = new System.Drawing.Point(304, 30);
			this.ChckBx_APT_FedEx.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.ChckBx_APT_FedEx.Name = "ChckBx_APT_FedEx";
			this.ChckBx_APT_FedEx.Size = new System.Drawing.Size(92, 20);
			this.ChckBx_APT_FedEx.TabIndex = 10;
			this.ChckBx_APT_FedEx.Text = "FedEx";
			this.ChckBx_APT_FedEx.UseVisualStyleBackColor = true;
			// 
			// ChckBx_APT_LH
			// 
			this.ChckBx_APT_LH.Location = new System.Drawing.Point(199, 27);
			this.ChckBx_APT_LH.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.ChckBx_APT_LH.Name = "ChckBx_APT_LH";
			this.ChckBx_APT_LH.Size = new System.Drawing.Size(107, 21);
			this.ChckBx_APT_LH.TabIndex = 9;
			this.ChckBx_APT_LH.Text = "Long Haul";
			this.ChckBx_APT_LH.UseVisualStyleBackColor = true;
			// 
			// TxtBox_APT_IATA
			// 
			this.TxtBox_APT_IATA.Location = new System.Drawing.Point(109, 27);
			this.TxtBox_APT_IATA.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.TxtBox_APT_IATA.Name = "TxtBox_APT_IATA";
			this.TxtBox_APT_IATA.Size = new System.Drawing.Size(72, 22);
			this.TxtBox_APT_IATA.TabIndex = 8;
			// 
			// TxtBox_APT_ICAO
			// 
			this.TxtBox_APT_ICAO.Location = new System.Drawing.Point(29, 27);
			this.TxtBox_APT_ICAO.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.TxtBox_APT_ICAO.Name = "TxtBox_APT_ICAO";
			this.TxtBox_APT_ICAO.Size = new System.Drawing.Size(65, 22);
			this.TxtBox_APT_ICAO.TabIndex = 7;
			//
			// AIP_SUP_report
			// 
			this.AIP_SUP_report.Controls.Add(this.Btn_Sup_printReport);
			this.AIP_SUP_report.Controls.Add(this.Btn_Sup_exportReport);
			this.AIP_SUP_report.Controls.Add(this.radBtn_Sup_31days);
			this.AIP_SUP_report.Controls.Add(this.radBtn_Sup_7days);
			this.AIP_SUP_report.Controls.Add(this.radBtn_Sup_24Hrs);
			this.AIP_SUP_report.Controls.Add(this.radBtn_Sup_noFilter);
			this.AIP_SUP_report.Controls.Add(this.btn_AIP_Sup_report);
			this.AIP_SUP_report.Controls.Add(this.Web_Sup_report);
			this.AIP_SUP_report.Location = new System.Drawing.Point(4, 25);
			this.AIP_SUP_report.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.AIP_SUP_report.Name = "AIP_SUP_report";
			this.AIP_SUP_report.Size = new System.Drawing.Size(1877, 1239);
			this.AIP_SUP_report.TabIndex = 7;
			this.AIP_SUP_report.Text = "AIP SUP Report";
			this.AIP_SUP_report.UseVisualStyleBackColor = true;
			// 
			// Btn_Sup_printReport
			// 
			this.Btn_Sup_printReport.BackColor = System.Drawing.Color.Black;
			this.Btn_Sup_printReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Btn_Sup_printReport.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
			this.Btn_Sup_printReport.Location = new System.Drawing.Point(715, 4);
			this.Btn_Sup_printReport.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Btn_Sup_printReport.Name = "Btn_Sup_printReport";
			this.Btn_Sup_printReport.Size = new System.Drawing.Size(100, 28);
			this.Btn_Sup_printReport.TabIndex = 15;
			this.Btn_Sup_printReport.Text = "Print";
			this.Btn_Sup_printReport.UseVisualStyleBackColor = false;
			this.Btn_Sup_printReport.Click += new System.EventHandler(this.Btn_Sup_printReportClick);
			//
			// Btn_Sup_exportReport
			//
			this.Btn_Sup_exportReport.BackColor = System.Drawing.Color.DarkGreen;
			this.Btn_Sup_exportReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Btn_Sup_exportReport.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
			this.Btn_Sup_exportReport.Location = new System.Drawing.Point(825, 4);
			this.Btn_Sup_exportReport.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Btn_Sup_exportReport.Name = "Btn_Sup_exportReport";
			this.Btn_Sup_exportReport.Size = new System.Drawing.Size(100, 28);
			this.Btn_Sup_exportReport.TabIndex = 16;
			this.Btn_Sup_exportReport.Text = "Export PDF";
			this.Btn_Sup_exportReport.UseVisualStyleBackColor = false;
			this.Btn_Sup_exportReport.Click += new System.EventHandler(this.Btn_Sup_exportReportClick);
			//
			// radBtn_Sup_31days
			// 
			this.radBtn_Sup_31days.Location = new System.Drawing.Point(568, 2);
			this.radBtn_Sup_31days.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_Sup_31days.Name = "radBtn_Sup_31days";
			this.radBtn_Sup_31days.Size = new System.Drawing.Size(139, 30);
			this.radBtn_Sup_31days.TabIndex = 14;
			this.radBtn_Sup_31days.Text = "31 days";
			this.radBtn_Sup_31days.UseVisualStyleBackColor = true;
			// 
			// radBtn_Sup_7days
			// 
			this.radBtn_Sup_7days.Location = new System.Drawing.Point(449, 2);
			this.radBtn_Sup_7days.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_Sup_7days.Name = "radBtn_Sup_7days";
			this.radBtn_Sup_7days.Size = new System.Drawing.Size(139, 30);
			this.radBtn_Sup_7days.TabIndex = 13;
			this.radBtn_Sup_7days.Text = "7 days";
			this.radBtn_Sup_7days.UseVisualStyleBackColor = true;
			// 
			// radBtn_Sup_24Hrs
			// 
			this.radBtn_Sup_24Hrs.Location = new System.Drawing.Point(284, 2);
			this.radBtn_Sup_24Hrs.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_Sup_24Hrs.Name = "radBtn_Sup_24Hrs";
			this.radBtn_Sup_24Hrs.Size = new System.Drawing.Size(139, 30);
			this.radBtn_Sup_24Hrs.TabIndex = 12;
			this.radBtn_Sup_24Hrs.Text = "Next 24Hrs";
			this.radBtn_Sup_24Hrs.UseVisualStyleBackColor = true;
			// 
			// radBtn_Sup_noFilter
			// 
			this.radBtn_Sup_noFilter.Checked = true;
			this.radBtn_Sup_noFilter.Location = new System.Drawing.Point(152, 2);
			this.radBtn_Sup_noFilter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.radBtn_Sup_noFilter.Name = "radBtn_Sup_noFilter";
			this.radBtn_Sup_noFilter.Size = new System.Drawing.Size(95, 30);
			this.radBtn_Sup_noFilter.TabIndex = 11;
			this.radBtn_Sup_noFilter.TabStop = true;
			this.radBtn_Sup_noFilter.Text = "No filter";
			this.radBtn_Sup_noFilter.UseVisualStyleBackColor = true;
			// 
			// btn_AIP_Sup_report
			// 
			this.btn_AIP_Sup_report.Location = new System.Drawing.Point(25, 2);
			this.btn_AIP_Sup_report.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.btn_AIP_Sup_report.Name = "btn_AIP_Sup_report";
			this.btn_AIP_Sup_report.Size = new System.Drawing.Size(100, 28);
			this.btn_AIP_Sup_report.TabIndex = 10;
			this.btn_AIP_Sup_report.Text = "Report !";
			this.btn_AIP_Sup_report.UseVisualStyleBackColor = true;
			this.btn_AIP_Sup_report.Click += new System.EventHandler(this.Btn_AIP_Sup_reportClick);
			// 
			// Web_Sup_report
			// 
			this.Web_Sup_report.Location = new System.Drawing.Point(3, 32);
			this.Web_Sup_report.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Web_Sup_report.MinimumSize = new System.Drawing.Size(20, 20);
			this.Web_Sup_report.Name = "Web_Sup_report";
			this.Web_Sup_report.Size = new System.Drawing.Size(1680, 1209);
			this.Web_Sup_report.TabIndex = 9;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1901, 1055);
			this.Controls.Add(this.tabControl1);
			this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Name = "MainForm";
			this.Text = "ICAO-CSV";
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.tabPage2.ResumeLayout(false);
			this.tabPage2.PerformLayout();
			this.tabPage3.ResumeLayout(false);
			this.tabPage4.ResumeLayout(false);
			this.tabPage5.ResumeLayout(false);
			this.APT_List.ResumeLayout(false);
			this.APT_List.PerformLayout();
			this.AIP_SUP_report.ResumeLayout(false);
			this.ResumeLayout(false);

		}
	}
}
