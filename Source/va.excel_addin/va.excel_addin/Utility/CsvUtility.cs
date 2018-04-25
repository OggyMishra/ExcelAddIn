namespace va.excel_addin_assignment.Utility.Implementation {
	using System;
	using System.IO;
	using System.Windows.Forms;
	using Microsoft.Office.Interop.Excel;

	internal static class CsvUtility {
		public static string FileName;
		private static bool IsCSVFilePresent(string fileName) {
			return System.IO.File.Exists(fileName);
		}

		internal static bool ImportCSV(Worksheet currentSheet, Range range, int[] columnDataTypes, bool autoFitColumns) {
			bool imported = false;
			ChooseFilePath();
			if (IsCSVFilePresent(FileName)) {
				currentSheet.QueryTables.Add(
					"TEXT;" + Path.GetFullPath(FileName),
			range, Type.Missing);
				currentSheet.QueryTables[1].Name = Path.GetFileNameWithoutExtension(FileName);
				currentSheet.QueryTables[1].FieldNames = true;
				currentSheet.QueryTables[1].RowNumbers = false;
				currentSheet.QueryTables[1].FillAdjacentFormulas = false;
				currentSheet.QueryTables[1].RefreshOnFileOpen = false;
				currentSheet.QueryTables[1].RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells;
				currentSheet.QueryTables[1].SavePassword = false;
				currentSheet.QueryTables[1].SaveData = true;
				currentSheet.QueryTables[1].AdjustColumnWidth = true;
				currentSheet.QueryTables[1].RefreshPeriod = 0;
				currentSheet.QueryTables[1].TextFilePromptOnRefresh = false;
				currentSheet.QueryTables[1].TextFilePlatform = 437;
				currentSheet.QueryTables[1].TextFileStartRow = 1;
				currentSheet.QueryTables[1].TextFileParseType = XlTextParsingType.xlDelimited;
				currentSheet.QueryTables[1].TextFileTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote;
				currentSheet.QueryTables[1].TextFileConsecutiveDelimiter = false;
				currentSheet.QueryTables[1].TextFileTabDelimiter = false;
				currentSheet.QueryTables[1].TextFileSemicolonDelimiter = false;
				currentSheet.QueryTables[1].TextFileCommaDelimiter = true;
				currentSheet.QueryTables[1].TextFileSpaceDelimiter = false;
				currentSheet.QueryTables[1].TextFileColumnDataTypes = columnDataTypes;
				currentSheet.QueryTables[1].Refresh(false);

				if (autoFitColumns == true)
					currentSheet.QueryTables[1].Destination.EntireColumn.AutoFit();

				imported = true;
			} else {
				MessageBox.Show("Not a valid path");
			}
			return imported;
		}

		private static void ChooseFilePath() {
			OpenFileDialog choofdlog = new OpenFileDialog();
			choofdlog.Filter = "All Files (*.csv*)|*.csv*";
			choofdlog.FilterIndex = 1;
			choofdlog.Multiselect = false;

			if (choofdlog.ShowDialog() == DialogResult.OK)
				FileName = choofdlog.FileName;
		}
	}
}
