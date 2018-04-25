namespace va.excel_addin_assignment.Styles {

	public static class Style {
		private static HeaderStyle _headerStyle;
		private static UpdatedItemStyle _updatedItemStyle;
		private static CellStyle _cellStyle;
		public static Microsoft.Office.Interop.Excel.Style Get(Microsoft.Office.Interop.Excel.Styles styles, StyleType type) {
			switch (type) {
				case StyleType.Header:
					if (_headerStyle == null) {
						_headerStyle = new HeaderStyle();
					}
					return _headerStyle.GetStyle(styles);
				case StyleType.UpdatedValue:
					if (_updatedItemStyle == null) {
						_updatedItemStyle = new UpdatedItemStyle();
					}
					return _updatedItemStyle.GetStyle(styles);
				default:
					if (_cellStyle == null) {
						_cellStyle = new CellStyle();
					}
					return _cellStyle.GetStyle(styles);
			}
		}
	}

	public enum StyleType {
		Cell = 1,
		UpdatedValue = 2,
		Header = 3
	}
}
