/*
 * Created by SharpDevelop.
 * User: jgazon
 * Date: 25-06-20
 * Time: 16:44
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Dispatch
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		private System.Windows.Forms.TabControl Tabs;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.WebBrowser Web_TAF;
		private System.Windows.Forms.Button Btn_addTAF;
		private System.Windows.Forms.RichTextBox RchTxt_TAF;
		private System.Windows.Forms.TabPage APT_List;
		private System.Windows.Forms.Button Btn_addAPT;
		private System.Windows.Forms.CheckBox ChckBx_APT_Charters;
		private System.Windows.Forms.CheckBox ChckBx_APT_FedEx;
		private System.Windows.Forms.CheckBox ChckBx_APT_LH;
		private System.Windows.Forms.TextBox TxtBox_APT_IATA;
		private System.Windows.Forms.TextBox TxtBox_APT_ICAO;
		private System.Windows.Forms.Button Btn_CopyAPTList;
		private System.Windows.Forms.Button Btn_CopyAPTList2;
		private System.Windows.Forms.Button Btn_refreshApp;
		private System.Windows.Forms.Button btn_editList;
		
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
			this.Tabs = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.Btn_CopyAPTList2 = new System.Windows.Forms.Button();
			this.Web_TAF = new System.Windows.Forms.WebBrowser();
			this.Btn_addTAF = new System.Windows.Forms.Button();
			this.RchTxt_TAF = new System.Windows.Forms.RichTextBox();
			this.APT_List = new System.Windows.Forms.TabPage();
			this.btn_editList = new System.Windows.Forms.Button();
			this.Btn_CopyAPTList = new System.Windows.Forms.Button();
			this.Btn_addAPT = new System.Windows.Forms.Button();
			this.ChckBx_APT_Charters = new System.Windows.Forms.CheckBox();
			this.ChckBx_APT_FedEx = new System.Windows.Forms.CheckBox();
			this.ChckBx_APT_LH = new System.Windows.Forms.CheckBox();
			this.TxtBox_APT_IATA = new System.Windows.Forms.TextBox();
			this.TxtBox_APT_ICAO = new System.Windows.Forms.TextBox();
			this.Btn_refreshApp = new System.Windows.Forms.Button();
			this.Tabs.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.APT_List.SuspendLayout();
			this.SuspendLayout();
			// 
			// Tabs
			// 
			this.Tabs.Controls.Add(this.tabPage1);
			this.Tabs.Controls.Add(this.APT_List);
			this.Tabs.Location = new System.Drawing.Point(66, 8);
			this.Tabs.Margin = new System.Windows.Forms.Padding(2);
			this.Tabs.Name = "Tabs";
			this.Tabs.SelectedIndex = 0;
			this.Tabs.Size = new System.Drawing.Size(1817, 977);
			this.Tabs.TabIndex = 0;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.Btn_CopyAPTList2);
			this.tabPage1.Controls.Add(this.Web_TAF);
			this.tabPage1.Controls.Add(this.Btn_addTAF);
			this.tabPage1.Controls.Add(this.RchTxt_TAF);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Margin = new System.Windows.Forms.Padding(2);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Padding = new System.Windows.Forms.Padding(2);
			this.tabPage1.Size = new System.Drawing.Size(1809, 951);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Weather";
			this.tabPage1.UseVisualStyleBackColor = true;
			// 
			// Btn_CopyAPTList2
			// 
			this.Btn_CopyAPTList2.Location = new System.Drawing.Point(1059, 50);
			this.Btn_CopyAPTList2.Margin = new System.Windows.Forms.Padding(2);
			this.Btn_CopyAPTList2.Name = "Btn_CopyAPTList2";
			this.Btn_CopyAPTList2.Size = new System.Drawing.Size(89, 23);
			this.Btn_CopyAPTList2.TabIndex = 6;
			this.Btn_CopyAPTList2.Text = "Copy APT List";
			this.Btn_CopyAPTList2.UseVisualStyleBackColor = true;
			this.Btn_CopyAPTList2.Click += new System.EventHandler(this.Btn_CopyAPTList2Click);
			// 
			// Web_TAF
			// 
			this.Web_TAF.Location = new System.Drawing.Point(22, 133);
			this.Web_TAF.Margin = new System.Windows.Forms.Padding(2);
			this.Web_TAF.MinimumSize = new System.Drawing.Size(14, 13);
			this.Web_TAF.Name = "Web_TAF";
			this.Web_TAF.Size = new System.Drawing.Size(1782, 815);
			this.Web_TAF.TabIndex = 5;
			// 
			// Btn_addTAF
			// 
			this.Btn_addTAF.Location = new System.Drawing.Point(1059, 15);
			this.Btn_addTAF.Margin = new System.Windows.Forms.Padding(2);
			this.Btn_addTAF.Name = "Btn_addTAF";
			this.Btn_addTAF.Size = new System.Drawing.Size(89, 23);
			this.Btn_addTAF.TabIndex = 4;
			this.Btn_addTAF.Text = "TAF Analysis";
			this.Btn_addTAF.UseVisualStyleBackColor = true;
			this.Btn_addTAF.Click += new System.EventHandler(this.Btn_addTAFClick);
			// 
			// RchTxt_TAF
			// 
			this.RchTxt_TAF.Location = new System.Drawing.Point(22, 15);
			this.RchTxt_TAF.Margin = new System.Windows.Forms.Padding(2);
			this.RchTxt_TAF.Name = "RchTxt_TAF";
			this.RchTxt_TAF.Size = new System.Drawing.Size(988, 114);
			this.RchTxt_TAF.TabIndex = 3;
			this.RchTxt_TAF.Text = "";
			// 
			// APT_List
			// 
			this.APT_List.Controls.Add(this.btn_editList);
			this.APT_List.Controls.Add(this.Btn_CopyAPTList);
			this.APT_List.Controls.Add(this.Btn_addAPT);
			this.APT_List.Controls.Add(this.ChckBx_APT_Charters);
			this.APT_List.Controls.Add(this.ChckBx_APT_FedEx);
			this.APT_List.Controls.Add(this.ChckBx_APT_LH);
			this.APT_List.Controls.Add(this.TxtBox_APT_IATA);
			this.APT_List.Controls.Add(this.TxtBox_APT_ICAO);
			this.APT_List.Location = new System.Drawing.Point(4, 22);
			this.APT_List.Margin = new System.Windows.Forms.Padding(2);
			this.APT_List.Name = "APT_List";
			this.APT_List.Size = new System.Drawing.Size(1809, 951);
			this.APT_List.TabIndex = 4;
			this.APT_List.Text = "APT List";
			this.APT_List.UseVisualStyleBackColor = true;
			// 
			// btn_editList
			// 
			this.btn_editList.Location = new System.Drawing.Point(19, 36);
			this.btn_editList.Margin = new System.Windows.Forms.Padding(2);
			this.btn_editList.Name = "btn_editList";
			this.btn_editList.Size = new System.Drawing.Size(75, 22);
			this.btn_editList.TabIndex = 7;
			this.btn_editList.Text = "Edit List";
			this.btn_editList.UseVisualStyleBackColor = true;
			this.btn_editList.Click += new System.EventHandler(this.Btn_editListClick);
			// 
			// Btn_CopyAPTList
			// 
			this.Btn_CopyAPTList.Location = new System.Drawing.Point(677, 5);
			this.Btn_CopyAPTList.Margin = new System.Windows.Forms.Padding(2);
			this.Btn_CopyAPTList.Name = "Btn_CopyAPTList";
			this.Btn_CopyAPTList.Size = new System.Drawing.Size(75, 22);
			this.Btn_CopyAPTList.TabIndex = 6;
			this.Btn_CopyAPTList.Text = "Copy List";
			this.Btn_CopyAPTList.UseVisualStyleBackColor = true;
			this.Btn_CopyAPTList.Click += new System.EventHandler(this.Btn_CopyAPTListClick);
			// 
			// Btn_addAPT
			// 
			this.Btn_addAPT.Location = new System.Drawing.Point(379, 5);
			this.Btn_addAPT.Margin = new System.Windows.Forms.Padding(2);
			this.Btn_addAPT.Name = "Btn_addAPT";
			this.Btn_addAPT.Size = new System.Drawing.Size(86, 22);
			this.Btn_addAPT.TabIndex = 5;
			this.Btn_addAPT.Text = "Add Airport !";
			this.Btn_addAPT.UseVisualStyleBackColor = true;
			this.Btn_addAPT.Click += new System.EventHandler(this.Btn_addAPTClick);
			// 
			// ChckBx_APT_Charters
			// 
			this.ChckBx_APT_Charters.Location = new System.Drawing.Point(290, 8);
			this.ChckBx_APT_Charters.Margin = new System.Windows.Forms.Padding(2);
			this.ChckBx_APT_Charters.Name = "ChckBx_APT_Charters";
			this.ChckBx_APT_Charters.Size = new System.Drawing.Size(69, 16);
			this.ChckBx_APT_Charters.TabIndex = 4;
			this.ChckBx_APT_Charters.Text = "Charters";
			this.ChckBx_APT_Charters.UseVisualStyleBackColor = true;
			// 
			// ChckBx_APT_FedEx
			// 
			this.ChckBx_APT_FedEx.Location = new System.Drawing.Point(225, 8);
			this.ChckBx_APT_FedEx.Margin = new System.Windows.Forms.Padding(2);
			this.ChckBx_APT_FedEx.Name = "ChckBx_APT_FedEx";
			this.ChckBx_APT_FedEx.Size = new System.Drawing.Size(69, 16);
			this.ChckBx_APT_FedEx.TabIndex = 3;
			this.ChckBx_APT_FedEx.Text = "FedEx";
			this.ChckBx_APT_FedEx.UseVisualStyleBackColor = true;
			// 
			// ChckBx_APT_LH
			// 
			this.ChckBx_APT_LH.Location = new System.Drawing.Point(146, 6);
			this.ChckBx_APT_LH.Margin = new System.Windows.Forms.Padding(2);
			this.ChckBx_APT_LH.Name = "ChckBx_APT_LH";
			this.ChckBx_APT_LH.Size = new System.Drawing.Size(80, 17);
			this.ChckBx_APT_LH.TabIndex = 2;
			this.ChckBx_APT_LH.Text = "Long Haul";
			this.ChckBx_APT_LH.UseVisualStyleBackColor = true;
			// 
			// TxtBox_APT_IATA
			// 
			this.TxtBox_APT_IATA.Location = new System.Drawing.Point(79, 6);
			this.TxtBox_APT_IATA.Margin = new System.Windows.Forms.Padding(2);
			this.TxtBox_APT_IATA.Name = "TxtBox_APT_IATA";
			this.TxtBox_APT_IATA.Size = new System.Drawing.Size(55, 20);
			this.TxtBox_APT_IATA.TabIndex = 1;
			// 
			// TxtBox_APT_ICAO
			// 
			this.TxtBox_APT_ICAO.Location = new System.Drawing.Point(19, 6);
			this.TxtBox_APT_ICAO.Margin = new System.Windows.Forms.Padding(2);
			this.TxtBox_APT_ICAO.Name = "TxtBox_APT_ICAO";
			this.TxtBox_APT_ICAO.Size = new System.Drawing.Size(50, 20);
			this.TxtBox_APT_ICAO.TabIndex = 0;
			// 
			// Btn_refreshApp
			// 
			this.Btn_refreshApp.Location = new System.Drawing.Point(3, 45);
			this.Btn_refreshApp.Name = "Btn_refreshApp";
			this.Btn_refreshApp.Size = new System.Drawing.Size(58, 43);
			this.Btn_refreshApp.TabIndex = 3;
			this.Btn_refreshApp.Text = "Refresh App";
			this.Btn_refreshApp.UseVisualStyleBackColor = true;
			this.Btn_refreshApp.Click += new System.EventHandler(this.Btn_refreshAppClick);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoScroll = true;
			this.AutoSize = true;
			this.ClientSize = new System.Drawing.Size(1431, 836);
			this.Controls.Add(this.Btn_refreshApp);
			this.Controls.Add(this.Tabs);
			this.Margin = new System.Windows.Forms.Padding(2);
			this.Name = "MainForm";
			this.Text = "Dispatch";
			this.Tabs.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.APT_List.ResumeLayout(false);
			this.APT_List.PerformLayout();
			this.ResumeLayout(false);

		}
	}
}
