using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelExport {
	class Program {
		private static bool ArrayTypeCheck(IRow nameRow, string nameToCheck, int ignoreColumnIndex) {
			if( nameToCheck.StartsWith("ignore") ) {
				return false;
			}
			for( int i = 0; i < nameRow.LastCellNum; i++ ) {
				if( ignoreColumnIndex == i ) {
					continue;
				}

				var cell = nameRow.GetCell(i);
				if( cell == null || cell.CellType == CellType.Blank ) {
					continue;
				}

				if( cell.StringCellValue == nameToCheck ) {
					return true;
				}
			}

			return false;
		}

		private static bool ClassTypeCheck(string nameToCheck, out string key, out string field) {
			key = string.Empty;
			field = string.Empty;
			if( nameToCheck.StartsWith("ignore") ) {
				return false;
			}

			if( nameToCheck.Contains(".") ) {
				var nameParts = nameToCheck.Split('.');
				key = nameParts[0];
				field = nameParts[1];
				return true;
			}
			return false;
		}

		private static void PressAnyKeyToContinue() {
			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine("Press any key to continue.");
			Console.ReadKey();
		}

		static void Main() {
			var excelFiles = JArray.Parse(File.ReadAllText("./Config/Export.json"));
			foreach( string excelFilePath in excelFiles ) {
				if( !ConvertExcelFile(excelFilePath) ) {
					PressAnyKeyToContinue();
					return;
				}
			}

			Console.WriteLine();
			Console.WriteLine();
			Console.WriteLine("Start server end export");
			m_exportScope = "s";
			m_serverFileAppend = "_Server";
			foreach( string excelFilePath in excelFiles ) {
				if(!ConvertExcelFile(excelFilePath)) {
					PressAnyKeyToContinue();
					return;
				}
			}

			var type = typeof(IPostExportCheck);
			var types = AppDomain.CurrentDomain.GetAssemblies()
				.SelectMany(s => s.GetTypes())
				.Where(p => type.IsAssignableFrom(p));

			foreach( var t in types ) {
				if( t.IsInterface ) {
					continue;
				}
				var checkInstance = Activator.CreateInstance(t) as IPostExportCheck;
				checkInstance.Check();
			}

			PressAnyKeyToContinue();
		}

		private static string m_exportScope = "c";
		private static string m_serverFileAppend = "";
		private static string m_exportDirectory = "../Export";

		private static bool ConvertExcelFile(string excelFilePath) {
			int currentExportRow = 0;
			var exportColumnName = "";
			var sheetName = "";
			Console.ForegroundColor = ConsoleColor.White;
			Console.WriteLine($"Open xlsx file: {excelFilePath}.");
			try {
				var i18nConfig = JObject.Parse(File.ReadAllText("Config/I18n.json"));
				var languageSupport = i18nConfig.GetValue("SupportLang") as JArray;

				var languages = new string[languageSupport.Count];
				var stringTypes = new string[languageSupport.Count];
				for (int l = 0; l < languageSupport.Count; l++) {
					languages[l] = languageSupport[l].ToString();
					stringTypes[l] = "string";
				}
				var keyRow = string.Join("\t", languages);
				var typeRowString = string.Join("\t", stringTypes);

				Dictionary<string, List<string>> translateText = new Dictionary<string, List<string>>(4096);
				foreach (var translateLine in File.ReadAllLines("../i18n/i18n.tsv")) {
					if (string.IsNullOrEmpty(translateLine)) {
						break;
					}
					var keyAndLanguageTexts = translateLine.Split('\t');
					var key = keyAndLanguageTexts[0];
					var texts = new List<string>(keyAndLanguageTexts.Skip(1));
					if (translateText.ContainsKey(key)) {
						Console.WriteLine("Exist Key : " + key);
						continue;
					}
					translateText.Add(key, texts);
				}
				File.Copy($"../{excelFilePath}", $"../{excelFilePath}.bak", true);
				excelFilePath = $"../{excelFilePath}.bak";

				using (var stream = new FileStream(excelFilePath, FileMode.Open)) {
					currentExportRow = 0;
					exportColumnName = "";
					XSSFWorkbook workbook = new XSSFWorkbook(stream);
					for (int i = 0; i < workbook.NumberOfSheets; i++) {
						sheetName = workbook.GetSheetName(i);
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
						if (structures.Count == 0) {
							continue;
						}
						string exportFilePath = $"{m_exportDirectory}/{sheetName}{m_serverFileAppend}.tsv";
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
							HashSet<string> idDuplicateCheck = new HashSet<string>();
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

								if (idDuplicateCheck.Contains(id)) {
									throw new Exception($"Id duplicate row ：{row.RowNum}, id : {id}");
								}
								else {
									idDuplicateCheck.Add(id);
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
							writer.WriteLine($"id\t{keyRow}");
							writer.WriteLine($"string\t{typeRowString}");

							Console.ForegroundColor = ConsoleColor.Yellow;
							foreach (var t in translates) {
								if (translateText.TryGetValue($"{sheetName}:{t.Key}", out var translate)) {
									var tranlatedLanguage = translate.Skip(1);
									writer.WriteLine($"{t.Key}\t{t.Value}\t{string.Join("\t", tranlatedLanguage)}");
									continue;
								}
								Console.WriteLine("Not found i18n key ：" + t.Key);
								writer.WriteLine($"{t.Key}\t{t.Value}");
							}
							Console.ForegroundColor = ConsoleColor.White;
						}
					}
				}
			}
			catch (Exception e) {
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine($"Error occurred while export sheet : {sheetName}, row : {currentExportRow}, column : {exportColumnName}, {e.Message}");
				Console.WriteLine(e.Message);
				return false;
			}
			finally {
				File.Delete(excelFilePath);
            }
			return true;
		}

		private static void BuildStructure(IRow nameRow, IRow typeRow, IRow clientServerScopeRow, List<int> exportColumnIndeies, out List<ITypeStructure> structures, out Dictionary<string, string> translates) {
			Dictionary<string, ComplexTypeStructure> complexTypes = new Dictionary<string, ComplexTypeStructure>();
			structures = new List<ITypeStructure>();
			translates = new Dictionary<string, string>();
			Console.WriteLine($"Export name rows : {string.Join(",", exportColumnIndeies)}");
			foreach( var index in exportColumnIndeies ) {
				var type = typeRow.GetCell(index);
				var name = nameRow.GetCell(index).StringCellValue;
				var csCell = clientServerScopeRow.GetCell(index);
				if( csCell  == null || csCell.CellType == CellType.Blank) {
					throw new Exception("C,S row is empty。");
				}
				var clientServerScope = clientServerScopeRow.GetCell(index).StringCellValue;
				if( !clientServerScope.Contains(m_exportScope) ) {
					continue;
				}

				ITypeStructure simpleType;
				var classType = false;
				if( ClassTypeCheck(name, out var key, out var field) ) {
					simpleType = TranslateSimpleType(translates, index, type.StringCellValue, field);
					classType = true;
				}
				else {
					simpleType = TranslateSimpleType(translates, index, type.StringCellValue, name);
					key = name;
				}

				// array type check
				if( complexTypes.TryGetValue(key, out var complexType) ) {
					complexType.AppendColumnIndex(simpleType);
				}
				else if( classType ) {
					complexType = new ClassTypeStructure(key);
					complexType.AppendColumnIndex(simpleType);
					complexTypes.Add(key, complexType);
					structures.Add(complexType);
				}
				else if( ArrayTypeCheck(nameRow, name, index) ) {
					if( type.StringCellValue.StartsWith("links") ) {
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

		private static ITypeStructure TranslateSimpleType(Dictionary<string, string> translates, int columnIndex, string cellValue, string columnName) {
			if( cellValue == "string" ) {
				return new StringTypeStructure(columnIndex, columnName);
			}
			else if( cellValue == "color" ) {
				return new ColorTypeStructure(columnIndex, columnName);
			}
			else if( cellValue == "json" ) {
				return new JsonTypeStructure(columnIndex, columnName);
			}	
			else if( cellValue == "int" || cellValue == "number" ) {
				return new NumberTypeStructure(columnIndex, columnName);
			}
			else if( cellValue == "bool" || cellValue == "boolean" ) {
				return new BooleanTypeStructure(columnIndex, columnName);
			}
			else if( cellValue == "translate" ) {
				return new TranslateTypeStructure(columnIndex, columnName, translates);
			}
			else {
				if( cellValue.StartsWith("link_") ) {
					var type = cellValue.Substring(5);
					var subTypeStructure = TranslateSimpleType(translates, columnIndex, type, columnName);
					return new LinkTypeStructure(columnIndex, columnName, (SimpleTypeStructure) subTypeStructure);
				}
				else if( cellValue.StartsWith("links_") ) {
					var type = cellValue.Substring(6);
					var subTypeStructure = TranslateSimpleType(translates, columnIndex, type, columnName);
					return new LinksTypeStructure(columnIndex, columnName, (SimpleTypeStructure) subTypeStructure);
				}
				throw new Exception("Unsupport type : " + cellValue);
			}
		}
	}
}
