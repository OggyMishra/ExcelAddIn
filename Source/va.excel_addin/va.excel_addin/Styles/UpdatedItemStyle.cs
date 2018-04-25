namespace va.excel_addin_assignment.Styles {
	public class UpdatedItemStyle : CustomStyle {
		private Microsoft.Office.Interop.Excel.Style _style;
		public override Microsoft.Office.Interop.Excel.Style GetStyle(Microsoft.Office.Interop.Excel.Styles styles) {
			if (_style == null) {
			  _style = styles.Add("UpdatedStyle");
				_style.NumberFormat = "[Blue]#,###.00;[Red](#,###.00);0";
				_style.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
			}
			return _style;
		}
	}
}
