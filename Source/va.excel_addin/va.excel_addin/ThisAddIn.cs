namespace va.excel_addin_assignment {
	using va.excel_addin_assignment.Common;
	public partial class ThisAddIn {
		private IDataLoader _dataLoader;
		private void ThisAddIn_Startup(object sender, System.EventArgs e) {
			AddInManager.Initialize();
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
			AddInManager.ShutDown();
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup() {
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
