namespace va.excel_addin_assignment {
	partial class VARibbonControl : Microsoft.Office.Tools.Ribbon.RibbonBase {
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public VARibbonControl()
				: base(Globals.Factory.GetRibbonFactory()) {
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			this.vaTab = this.Factory.CreateRibbonTab();
			this.dataGrp = this.Factory.CreateRibbonGroup();
			this.loadBtn = this.Factory.CreateRibbonButton();
			this.separator1 = this.Factory.CreateRibbonSeparator();
			this.refreshBtn = this.Factory.CreateRibbonButton();
			this.vaTab.SuspendLayout();
			this.dataGrp.SuspendLayout();
			this.SuspendLayout();
			// 
			// vaTab
			// 
			this.vaTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.vaTab.Groups.Add(this.dataGrp);
			this.vaTab.Label = "Visible Alpha";
			this.vaTab.Name = "vaTab";
			// 
			// dataGrp
			// 
			this.dataGrp.Items.Add(this.loadBtn);
			this.dataGrp.Items.Add(this.separator1);
			this.dataGrp.Items.Add(this.refreshBtn);
			this.dataGrp.Label = "Data";
			this.dataGrp.Name = "dataGrp";
			// 
			// loadBtn
			// 
			this.loadBtn.Image = global::va.excel_addin_assignment.Properties.Resources.data_copy__1_;
			this.loadBtn.Label = "Load Data";
			this.loadBtn.Name = "loadBtn";
			this.loadBtn.ScreenTip = "Load data from external csv file";
			this.loadBtn.ShowImage = true;
			this.loadBtn.SuperTip = "Load Data";
			this.loadBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadBtn_Click);
			// 
			// separator1
			// 
			this.separator1.Name = "separator1";
			// 
			// refreshBtn
			// 
			this.refreshBtn.Image = global::va.excel_addin_assignment.Properties.Resources.cloud_computing;
			this.refreshBtn.Label = "Refresh Data";
			this.refreshBtn.Name = "refreshBtn";
			this.refreshBtn.ScreenTip = "Refresh data to external csv file";
			this.refreshBtn.ShowImage = true;
			this.refreshBtn.SuperTip = "Refresh Data";
			this.refreshBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refreshBtn_Click);
			// 
			// VARibbonControl
			// 
			this.Name = "VARibbonControl";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.vaTab);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.VARibbonControl_Load);
			this.vaTab.ResumeLayout(false);
			this.vaTab.PerformLayout();
			this.dataGrp.ResumeLayout(false);
			this.dataGrp.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab vaTab;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup dataGrp;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton loadBtn;
		internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshBtn;
	}

	partial class ThisRibbonCollection {
		internal VARibbonControl VARibbonControl
		{
			get { return this.GetRibbon<VARibbonControl>(); }
		}
	}
}
