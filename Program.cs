using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelExport {
	class Program {
		private static bool ArrayTypeCheck(IRow nameRow, string nameToCheck, int ignoreColumnIndex) {
			if (nameToCheck.StartsWith("ignore")) {
				return false;
			}
			for (int i = 0; i < nameRow.LastCellNum; i++) {
				if (ignoreColumnIndex == i) {
					continue;
				}

				var cell = nameRow.GetCell(i);
				if (cell == null || cell.CellType == CellType.Blank) {
					continue;
				}

				if (cell.StringCellValue == nameToCheck) {
					return true;
				}
			}

			return false;
		}

		private static bool ClassTypeCheck(string nameToCheck, out string key, out string field) {
			key = string.Empty;
			field = string.Empty;
			if (nameToCheck.StartsWith("ignore")) {
				return false;
			}

			if (nameToCheck.Contains(".")) {
				var nameParts = nameToCheck.Split('.');
				key = nameParts[0];
				field = nameParts[1];
				return true;
			}
			return false;
		}

		static void Main() {
			var excelFiles = JArray.Parse(File.ReadAllText("./Config/Export.json"));
			foreach (string excelFilePath in excelFiles) {
				ConvertExcelFile(excelFilePath);
			}

			Console.WriteLine();
			Console.WriteLine();
			Console.WriteLine("Start server end export");
			m_exportScope = "s";
			m_serverFileAppend = "_Server";
			foreach (string excelFilePath in excelFiles) {
				ConvertExcelFile(excelFilePath);
			}
			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine("Press any key to continue.");
			Console.ReadKey();
		}

		private static string m_exportScope = "c";
		private static string m_serverFileAppend = "";

		private static void ConvertExcelFile(string excelFilePath) {
			int currentExportRow = 0;
			var exportColumnName = "";
			Console.ForegroundColor = ConsoleColor.White;
			Console.WriteLine($"Open xlsx file: {excelFilePath}.");
			try {
				using (var stream = new FileStream($"../{excelFilePath}", FileMode.Open)) {
					XSSFWorkbook workbook = new XSSFWorkbook(stream);
					for (int i = 0; i < workbook.NumberOfSheets; i++) {
						var sheetName = workbook.GetSheetName(i);
						if (sheetName.StartsWith("ignore", StringComparison.InvariantCultureIgnoreCase)) {
							continue;
						}
						Console.WriteLine($"Export sheet : {sheetName}");
						var sheet = workbook.GetSheetAt(i);

						var nameRow = sheet.GetRow(0);
						var typeRow = sheet.GetRow(1);
						var clientServerScopeRow = sheet.GetRow(3);

						List<int> exportColumnIndeies = new List<int>();
						for (int columnIndex = 0; columnIndex < nameRow.LastCellNum; columnIndex++) {
							var cell = nameRow.GetCell(columnIndex);
							if (cell == null || cell.CellType == CellType.Blank) {
								continue;
							}

							if (cell.CellType == CellType.String) {
								if (cell.StringCellValue.StartsWith("ignore", StringComparison.InvariantCultureIgnoreCase)) {
									continue;
								}

								exportColumnIndeies.Add(columnIndex);
							}
						}

						BuildStructure(nameRow, typeRow, clientServerScopeRow, exportColumnIndeies, out var structures, out var translates);
						if(structures.Count == 0) {
							continue;
						}
						string exportFilePath = $"../Export/{sheetName}{m_serverFileAppend}.tsv";
						Console.WriteLine($"Export {sheetName} to {exportFilePath} contains row : {sheet.LastRowNum - 2}");
						using (var fileStream = new FileStream(exportFilePath, FileMode.Create))
						using (var writer = new StreamWriter(fileStream)) {
							string[] names = new string[structures.Count];
							string[] types = new string[structures.Count];
							for (int structureIndex = 0; structureIndex < structures.Count; structureIndex++) {
								var structure = structures[structureIndex];
								names[structureIndex] = structure.ColumnName;
								types[structureIndex] = structure.ColumnType;
							}

							writer.WriteLine(string.Join("\t", names));
							writer.WriteLine(string.Join("\t", types));

							string[] values = new string[structures.Count];
							int ignoreHeaderRowCount = 0;
							foreach (IRow row in sheet) {
								if (ignoreHeaderRowCount < 4) {
									ignoreHeaderRowCount++;
									continue;
								}
								if (row == null || row.Cells.Count == 0) {
									continue;
								}
								currentExportRow = row.RowNum;
								var id = structures[0].ConvertValue(row);
								if (string.IsNullOrEmpty(id)) {
									continue;
								}
								for (int structureIndex = 0; structureIndex < structures.Count; structureIndex++) {
									var structure = structures[structureIndex];
									if (structure is TranslateTypeStructure) {
										var translateStructure = structure as TranslateTypeStructure;
										translateStructure.ID = id;
									}
									exportColumnName = structure.ColumnName;
									values[structureIndex] = structure.ConvertValue(row);
								}
								writer.WriteLine(string.Join("\t", values));
							}
						}

						if (translates.Count == 0) {
							continue;
						}

						string exportI18nFilePath = $"../Export/{sheetName}{m_serverFileAppend}_i18n.tsv";
						using (var fileStream = new FileStream(exportI18nFilePath, FileMode.Create))
						using (var writer = new StreamWriter(fileStream)) {
							writer.WriteLine("id\tcn");
							writer.WriteLine("string\tstring");

							foreach (var t in translates) {
								writer.WriteLine($"{t.Key}\t{t.Value}");
							}
						}
					}
				}

				IPostExportCheck check = new ExtendCheck();
				check.Check();
			}
			catch (Exception e) {
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine($"Error occurred while export row : {currentExportRow}, {exportColumnName}");
				Console.WriteLine(e.Message);
			}
		}

		private static void BuildStructure(IRow nameRow, IRow typeRow, IRow clientServerScopeRow, List<int> exportColumnIndeies, out List<ITypeStructure> structures, out Dictionary<string, string> translates) {
			Dictionary<string, ComplexTypeStructure> complexTypes = new Dictionary<string, ComplexTypeStructure>();
			structures = new List<ITypeStructure>();
			translates = new Dictionary<string, string>();
			Console.WriteLine($"Export name rows : {string.Join(",", exportColumnIndeies)}");
			foreach (var index in exportColumnIndeies) {
				var type = typeRow.GetCell(index);
				var name = nameRow.GetCell(index).StringCellValue;
				var clientServerScope = clientServerScopeRow.GetCell(index).StringCellValue;
				if(!clientServerScope.Contains(m_exportScope)) {
					continue;
				}

				ITypeStructure simpleType;
				var classType = false;
				if (ClassTypeCheck(name, out var key, out var field)) {
					simpleType = TranslateSimpleType(translates, index, type, field);
					classType = true;
				}
				else {
					simpleType = TranslateSimpleType(translates, index, type, name);
					key = name;
				}

				// array type check
				if (complexTypes.TryGetValue(key, out var complexType)) {
					complexType.AppendColumnIndex(simpleType);
				}
				else if (classType) {
					complexType = new ClassTypeStructure(key);
					complexType.AppendColumnIndex(simpleType);
					complexTypes.Add(key, complexType);
					structures.Add(complexType);
				}
				else if (ArrayTypeCheck(nameRow, name, index)) {
					if (type.StringCellValue == "link") {
						complexType = new JsonArrayTypeStructure(name);
					}
					else {
						complexType = new ArrayTypeStructure(name);
					}

					complexType.AppendColumnIndex(simpleType);
					complexTypes.Add(name, complexType);
					structures.Add(complexType);
				}
				else {
					structures.Add(simpleType);
				}
			}
		}

		private static ITypeStructure TranslateSimpleType(Dictionary<string, string> translates, int columnIndex, ICell typeCell, string columnName) {
			if (typeCell.StringCellValue == "string") {
				return new StringTypeStructure(columnIndex, columnName);
			}
			else if (typeCell.StringCellValue == "color") {
				return new ColorTypeStructure(columnIndex, columnName);
			}
			else if (typeCell.StringCellValue == "int" || typeCell.StringCellValue == "number") {
				return new NumberTypeStructure(columnIndex, columnName);
			}
			else if (typeCell.StringCellValue == "bool" || typeCell.StringCellValue == "boolean") {
				return new BooleanTypeStructure(columnIndex, columnName);
			}
			else if (typeCell.StringCellValue == "translate") {
				return new TranslateTypeStructure(columnIndex, columnName, translates);
			}
			else if (typeCell.StringCellValue == "link") {
				return new LinkTypeStructure(columnIndex, columnName);
			}
			else if (typeCell.StringCellValue == "links") {
				return new LinksTypeStructure(columnIndex, columnName);
			}
			else {
				throw new Exception("Unsupport type : " + typeCell.StringCellValue);
			}
		}
	}
}
