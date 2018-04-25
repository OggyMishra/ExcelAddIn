namespace va.excel_addin_assignment.Styles {
	public class CellStyle : CustomStyle {
		private Microsoft.Office.Interop.Excel.Style _style;
		public override Microsoft.Office.Interop.Excel.Style GetStyle(Microsoft.Office.Interop.Excel.Styles styles) {
			if (_style == null) {
				_style = styles.Add("CellStyle");
				_style.NumberFormat = "[Black]#,###.00;[Red](#,###.00);0";
			}
			return _style;
		}
	}
}
