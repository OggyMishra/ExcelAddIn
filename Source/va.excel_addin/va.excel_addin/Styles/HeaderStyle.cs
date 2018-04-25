namespace va.excel_addin_assignment.Styles {
	public class HeaderStyle : CustomStyle {
		private Microsoft.Office.Interop.Excel.Style _style;
		public override Microsoft.Office.Interop.Excel.Style GetStyle(Microsoft.Office.Interop.Excel.Styles styles) {
			if (_style == null) {
				_style = styles.Add("HeaderStyle");
				_style.Font.Bold = true;
			}
			return _style;
		}
	}
}
