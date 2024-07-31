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
		public abstract JToken ConvertJson(IRow row);
	}


	public class ColorTypeStructure : SimpleTypeStructure {
		public ColorTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "color";

		public override JToken ConvertJson(IRow row) {
			var val = ConvertValue(row);
			if (string.IsNullOrEmpty(val)) {
				return null;
			}
			return val;
		}

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return null;
			}
			return cell.StringCellValue;
		}
	}

	public class JsonTypeStructure : SimpleTypeStructure {
		public JsonTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}

		public override string ColumnType => "json";

		public override JToken ConvertJson(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return null;
			}

			return cell.StringCellValue;
		}

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return string.Empty;
			}

			return cell.StringCellValue;
		}
	}

	public class StringTypeStructure : SimpleTypeStructure {
		public StringTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}

		public override string ColumnType => "string";

		public override JToken ConvertJson(IRow row) {
			var strValue = ConvertValue(row);
			if (string.IsNullOrEmpty(strValue)) {
				return null;
			}
			return strValue;
		}

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return string.Empty;
			}

			var ret = string.Empty;
			if (cell.CellType == CellType.Numeric) {
				ret = cell.NumericCellValue.ToString();
			}
			else if (cell.CellType == CellType.Formula) {
				if (cell.CachedFormulaResultType == CellType.String) {
					ret = cell.StringCellValue;
				}
				else if (cell.CachedFormulaResultType == CellType.Numeric) {
					ret =  cell.NumericCellValue.ToString();
				}
				else {
					ret = cell.StringCellValue;
				}
			}
			else if (cell.CellType == CellType.String) {
				ret = cell.StringCellValue;
			}
			else {
				ret = cell.ToString();
			}

			ret = ret.Trim();
			return ret;
		}
	}

	public class NumberTypeStructure : SimpleTypeStructure {
		public NumberTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "number";

		public override JToken ConvertJson(IRow row) {
			var strValue = ConvertValue(row);
			if (string.IsNullOrEmpty(strValue)) {
				return null;
			}

			if (float.TryParse(strValue, out var val)) {
				return val;
			}
			var cell = row.GetCell(m_columnIndex);
			return cell.NumericCellValue;
		}

		public override string ConvertValue(IRow row) {
			var cell = row.GetCell(m_columnIndex);
			if (cell == null || cell.CellType == CellType.Blank) {
				return string.Empty;
			}

			if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.String) {
				return cell.StringCellValue;
			}

			if (cell.CellType == CellType.String) {
				if (!float.TryParse(cell.StringCellValue, out var val)) {
					throw new System.Exception($"Cell value {cell.StringCellValue} can not convert to number");
				}
				return cell.StringCellValue;
			}
			return cell.NumericCellValue.ToString();
		}
	}

	public class BooleanTypeStructure : SimpleTypeStructure {
		public BooleanTypeStructure(int columnIndex, string columnName) : base(columnIndex, columnName) {
		}
		public override string ColumnType => "boolean";

		public override JToken ConvertJson(IRow row) {
			return ConvertValue(row) == "1";
		}

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
		private SimpleTypeStructure m_subTypeStructure;
		public LinkTypeStructure(int columnIndex, string columnName, SimpleTypeStructure typeStructure) : base(columnIndex, columnName) {
			m_subTypeStructure = typeStructure;
		}
		public override string ColumnType => "link";

		public override JToken ConvertJson(IRow row) {
			return m_subTypeStructure.ConvertJson(row);
		}

		public override string ConvertValue(IRow row) {
			return m_subTypeStructure.ConvertValue(row);
		}
	}
	public class LinksTypeStructure : SimpleTypeStructure {
		private SimpleTypeStructure m_subTypeStructure;
		public LinksTypeStructure(int columnIndex, string columnName, SimpleTypeStructure typeStructure) : base(columnIndex, columnName) {
			m_subTypeStructure = typeStructure;
		}
		public override string ColumnType => "links";

		public override JToken ConvertJson(IRow row) {
			return m_subTypeStructure.ConvertJson(row);
		}

		public override string ConvertValue(IRow row) {
			return m_subTypeStructure.ConvertValue(row);
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

		public override JToken ConvertJson(IRow row) {
			throw new System.Exception("translate type can not be included in a complex type.");
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
				var simpleType = columnIndex as SimpleTypeStructure;
				var value = simpleType.ConvertJson(row);
				if (value == null) {
					continue;
				}
				jobject.Add(columnIndex.ColumnName, value);
			}

			if (!jobject.HasValues) {
				return "";
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

			var array = new JArray();
			foreach (var columnIndex in m_columnIndeise) {
				var simpleType = columnIndex as SimpleTypeStructure;
				var val = simpleType.ConvertJson(row);
				if (val == null) {
					continue;
				}
				array.Add(val);
			}

			return array.ToString(Newtonsoft.Json.Formatting.None);
		}
	}

	public class JsonArrayTypeStructure : ArrayTypeStructure {
		public JsonArrayTypeStructure(string columnName) : base(columnName) {

		}

		public override string ColumnType => "links";
	}
}
