using Newtonsoft.Json.Linq;
using System.IO;

namespace ExcelExport {
	internal static class ConfigUtil {
		private static JToken ProcessConfigColumnToJson(string type, string val) {
			switch (type) {
				case "int":
					return int.Parse(val);
				case "string":
					return val;
				case "number":
					return float.Parse(val);
				case "boolean":
				case "bool":
					return val == "true" || val == "1";
				case "json":
					return JToken.Parse(val);
			}

			return null;
		}

		public static JObject LoadConfigToJson(string filename) {
			var json = new JObject();
			var path = filename + ".tsv";
			using( var reader = new StringReader(File.ReadAllText(path)) ) {
				var keys = reader.ReadLine();
				var types = reader.ReadLine();

				var keyArr = keys.Split('\t');
				var typeArr = types.Split('\t');

				while (true) {
					var row = reader.ReadLine();
					if (string.IsNullOrEmpty(row)) {
						break;
					}

					var rowDataArr = row.Split('\t');

					var rowJson = new JObject();
					for (int i = 0; i < rowDataArr.Length; i++) {
						rowJson[keyArr[i]] = ProcessConfigColumnToJson(typeArr[i], rowDataArr[i]);
					}

					var id = rowDataArr[0];
					json[id] = rowJson;
				}
			}

			return json;
		}
	}
}
