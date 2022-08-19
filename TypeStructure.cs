using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExcelExport {
	public interface ITypeStructure {
		string ColumnName { get; }

		string ColumnType { get; }

		string ConvertValue(IRow row);
	}

	public abstract class SimpleTypeStructure : ITypeStructure {
		protected readonly int m_columnIndex;
		protected readonly string m_columnName;
		public string ColumnName => m_columnName;

		public abstract string ColumnType { get; }

		public SimpleTypeStructure(int columnIndex, string columnName) {
			m_columnIndex = columnIndex;
			m_columnName = columnName;
		}

		public abstract string ConvertValue(IRow row);
	}


	public class ColorTypeStructure : SimpleTypeStructure {
		public ColorTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "color";
		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			return cell.StringCellValue;
		}
	}

	public class StringTypeStructure : SimpleTypeStructure {
		public StringTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}

		public override string ColumnType => "string";

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return string.Empty;
			}

			if (cell.CellType == CellType.Numeric) {
				return cell.NumericCellValue.ToString();
			}
			else if (cell.CellType == CellType.Formula) {
				if (cell.CachedFormulaResultType == CellType.String) {
					return cell.StringCellValue;
				}
				else if (cell.CachedFormulaResultType == CellType.Numeric) {
					return cell.NumericCellValue.ToString();
				}
				else {
					return cell.StringCellValue;
				}
			}
			else if (cell.CellType == CellType.String) {
				return cell.StringCellValue;
			}
			return cell.ToString();
		}
	}

	public class NumberTypeStructure : SimpleTypeStructure {
		public NumberTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "number";

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return string.Empty;
			}

			return cell.NumericCellValue.ToString();
		}
	}

	public class BooleanTypeStructure : SimpleTypeStructure {
		public BooleanTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "boolean";

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return "0";
			}

			if (cell.CellType == CellType.Boolean) {
				return cell.BooleanCellValue ? "1" : "0";
			}
			else if (cell.CellType == CellType.Numeric) {
				return cell.NumericCellValue == 1 ? "1" : "0";
			}
			else {
				return cell.StringCellValue == "true" ? "1" : "0";
			}
		}
	}

	public class LinkTypeStructure : SimpleTypeStructure {
		public LinkTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "link";

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return "";
			}

			if (cell.CellType == CellType.String || cell.CellType == CellType.Formula) {
				return cell.StringCellValue;
			}
			else {
				return cell.NumericCellValue.ToString();
			}
		}
	}
	public class LinksTypeStructure : SimpleTypeStructure {
		public LinksTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "links";

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return "";
			}

			if (cell.CellType == CellType.String || cell.CellType == CellType.Formula) {
				return cell.StringCellValue;
			}
			else {
				return cell.NumericCellValue.ToString();
			}
		}
	}

	public class TranslateTypeStructure : SimpleTypeStructure {
		private readonly Dictionary<string, string> m_translates;

		public string ID { private get; set; }
		public override string ColumnType => "translate";

		public TranslateTypeStructure(int columnIndex, string columnName, Dictionary<string, string> translates) : base(columnIndex, columnName) {
			m_translates = translates;
		}

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return "";
			}

			if (cell.CellType == CellType.String) {
				m_translates.Add($"{m_columnName}:{ID}", cell.StringCellValue);
			}
			return "";
		}
	}

	public abstract class ComplexTypeStructure : ITypeStructure {
		protected readonly List<ITypeStructure> m_columnIndeise = new List<ITypeStructure>();
		protected readonly string m_columnName;
		public ComplexTypeStructure(string columnName) {
			m_columnName = columnName;
		}

		public void AppendColumnIndex(ITypeStructure columnType) {
			m_columnIndeise.Add(columnType);
		}

		public string ColumnName => m_columnName;
		public abstract string ColumnType { get; }
		public abstract string ConvertValue(IRow row);
	}


	public class ClassTypeStructure : ComplexTypeStructure {
		public ClassTypeStructure(string columnName) : base(columnName) {
		}

		public override string ColumnType => "json";

		public override string ConvertValue(IRow row) {
			if (m_columnIndeise.Count == 0) {
				throw new System.Exception("No valid columns");
			}

			JObject jobject = new JObject();
			foreach (var columnIndex in m_columnIndeise) {
				jobject.Add(columnIndex.ColumnName, columnIndex.ConvertValue(row));
			}

			return jobject.ToString(Newtonsoft.Json.Formatting.None);
		}
	}

	public class ArrayTypeStructure : ComplexTypeStructure {
		public ArrayTypeStructure(string columnName) : base(columnName) {

		}
		public override string ColumnType => "json";
		public override string ConvertValue(IRow row) {
			if (m_columnIndeise.Count == 0) {
				throw new System.Exception("No valid columns");
			}

			List<string> values = new List<string>();
			foreach (var columnIndex in m_columnIndeise) {
				var val = columnIndex.ConvertValue(row);
				if (string.IsNullOrEmpty(val)) {
					continue;
				}
				values.Add(val);
			}

			return new JArray(values).ToString(Newtonsoft.Json.Formatting.None);
		}
	}

	public class JsonArrayTypeStructure : ArrayTypeStructure {
		public JsonArrayTypeStructure(string columnName) : base(columnName) {

		}

		public override string ColumnType => "links";
	}
}
