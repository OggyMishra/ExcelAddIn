namespace va.excel_addin_assignment.Styles {
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Text;
	using System.Threading.Tasks;

	public abstract class CustomStyle {
		public abstract Microsoft.Office.Interop.Excel.Style GetStyle(Microsoft.Office.Interop.Excel.Styles styles);
	}
}
