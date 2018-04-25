namespace va.excel_addin_assignment.Common {
	using Utility.Implementation;
	using Interop = Microsoft.Office.Interop.Excel;

	public class DataLoader : IDataLoader {
		public void LoadData(Interop.Workbook workBook) {
			var workSheet = (Interop.Worksheet)(workBook.ActiveSheet);
			var rgDesc = "A1:Z1";
			var range = workSheet.Range[rgDesc];
			if (CsvUtility.ImportCSV(workSheet, range, new int[] { 2 }, true)) {
				workSheet.Change += WorkSheet_Change;
				this.SetCustomStyle(workBook);
			}
		}

		private void WorkSheet_Change(Interop.Range target) {
			var wb = ((target.Parent as Interop.Worksheet).Parent as Interop.Workbook);
			StyleSetter.SetModifiedCellStyle(wb, target);
		}

		public void RefreshData(Interop.Workbook workBook) {
			workBook.SaveAs(CsvUtility.FileName, Interop.XlFileFormat.xlCSV);
			StyleSetter.SetCellStyle(workBook);
			StyleSetter.SetHeaderStyle(workBook);
		}

		private void SetCustomStyle(Interop.Workbook workBook) {
			StyleSetter.SetCellStyle(workBook);
			StyleSetter.SetHeaderStyle(workBook);
		}
	}
	public interface IDataLoader {
		void LoadData(Interop.Workbook workBook);
		void RefreshData(Interop.Workbook workBook);
	}
}
