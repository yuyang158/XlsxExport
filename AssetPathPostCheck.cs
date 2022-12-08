using Newtonsoft.Json.Linq;
using System;
using System.IO;

namespace ExcelExport {
	internal class AssetPathPostCheck : IPostExportCheck {
		private readonly JObject m_assetPathCheckConfig;
		private const string m_assetPathPrefix = "../../XiaoiceIslandTiledMap/";

		public AssetPathPostCheck() {
			m_assetPathCheckConfig = JObject.Parse(File.ReadAllText("Config/AssetPath.json"));
		}

		private void TestAssetPathExist(string assetPrefix, string assetPath) {
			var p = m_assetPathPrefix + assetPrefix + assetPath;
			p = Path.GetFullPath(p);

			if (File.Exists(p)) {
				var fileInfo = new FileInfo(p);
				if (fileInfo.FullName == p) {
					return;
				}
			}
			Console.WriteLine($"ERROR: Asset Not Exist : {assetPath}!");
		}

		public void Check() {
			Console.ForegroundColor = ConsoleColor.Red;
			foreach (var item in m_assetPathCheckConfig) {
				var config = ConfigUtil.LoadConfigToJson($"../Export/{item.Key}");
				var checkContent = item.Value as JObject;

				var columnNameArray = checkContent["Columns"] as JArray;
				var assetPrefix = checkContent["Prefix"].ToString();
				foreach (var columnName in columnNameArray) {
					foreach (var configPair in config) {
						var configRow = configPair.Value as JObject;
						var cellContent = configRow[columnName.ToString()];
						if (cellContent.Type == JTokenType.String) {
							var assetPath = cellContent.ToString();
							TestAssetPathExist(assetPrefix, assetPath);
						}
						else if (cellContent.Type == JTokenType.Array) {
							foreach (var assetPath in cellContent) {
								TestAssetPathExist(assetPrefix, assetPath.ToString());
							}
						}
					}
				}
			}
		}
	}
}
