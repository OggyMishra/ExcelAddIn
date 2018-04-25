namespace va.excel_addin_assignment.Utility.Implementation {
	using Styles;

	/// <summary>
	/// Style setting utility
	/// </summary>
	public static class StyleSetter {

		#region Exposed methods
		/// <summary>
		/// Sets the header style.
		/// </summary>
		/// <param name="workbook">The workbook.</param>
		internal static void SetHeaderStyle(Microsoft.Office.Interop.Excel.Workbook workbook) {
			Microsoft.Office.Interop.Excel.Range rows = workbook.ActiveSheet.Rows;
			rows[1].Style = Style.Get(workbook.Styles, StyleType.Header);

			Microsoft.Office.Interop.Excel.Range cols = workbook.ActiveSheet.Columns;
			cols[1].Style = Style.Get(workbook.Styles, StyleType.Header);
		}
		/// <summary>
		/// Sets the cell style.
		/// </summary>
		/// <param name="workbook">The workbook.</param>
		internal static void SetCellStyle(Microsoft.Office.Interop.Excel.Workbook workbook) {
			SetStyle(workbook, StyleType.Cell);
		}

		/// <summary>
		/// Sets the modified cell style.
		/// </summary>
		/// <param name="workbook">The workbook.</param>
		internal static void SetModifiedCellStyle(Microsoft.Office.Interop.Excel.Workbook workbook, Microsoft.Office.Interop.Excel.Range cell) {
			cell.Style = Style.Get(workbook.Styles, StyleType.UpdatedValue);
		}
		#endregion

		/// <summary>
		/// Sets the style.
		/// </summary>
		/// <param name="workbook">The workbook.</param>
		/// <param name="type">The type.</param>
		private static void SetStyle(Microsoft.Office.Interop.Excel.Workbook workbook, StyleType type) {
			Microsoft.Office.Interop.Excel.Range cells = workbook.ActiveSheet.Cells;
			cells.Style = Style.Get(workbook.Styles, type);
		}
	}
}
