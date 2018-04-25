namespace va.excel_addin_assignment {
	using Microsoft.Office.Tools.Ribbon;
	using Microsoft.Practices.Unity;
	using va.excel_addin_assignment.Common;

	public partial class VARibbonControl {
		private IDataLoader _dataLoader;
		private void VARibbonControl_Load(object sender, RibbonUIEventArgs e) {
			_dataLoader = AddInManager.GetUnityInstance().Resolve<IDataLoader>();
		}

		private void loadBtn_Click(object sender, RibbonControlEventArgs e) {
			var workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
			_dataLoader.LoadData(workBook);
		}

		private void refreshBtn_Click(object sender, RibbonControlEventArgs e) {
			var workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
			_dataLoader.RefreshData(workBook);
		}
	}
}
