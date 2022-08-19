using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelExport {
	public interface IPostExportCheck {
		void Check();
	}

	public class ExtendCheck : IPostExportCheck {
		private readonly JObject m_extendRelations;

		public ExtendCheck() {
			m_extendRelations = JObject.Parse(File.ReadAllText("./Config/ExtendRelation.json"));
		}

		private static List<string> ReadTsvKeys(string tsvName) {
			var lines = File.ReadAllLines($"../Export/{tsvName}.tsv");
			var keys = new List<string>();
			for (int i = 3; i < lines.Length; i++) {
				if(string.IsNullOrEmpty(lines[i])) {
					continue;
				}
				var lineParts = lines[i].Split('\t');
				keys.Add(lineParts[0]);
			}
			return keys;
		}

		public void Check() {
			Console.ForegroundColor = ConsoleColor.DarkRed;
			foreach (var relation in m_extendRelations) {
				var baseTsvKeys = ReadTsvKeys(relation.Key.ToString());
				foreach (var extendTsvName in relation.Value) {
					var extendTsvKeys = ReadTsvKeys(extendTsvName.ToString());
					foreach (var key in extendTsvKeys) {
						if(!baseTsvKeys.Contains(key)) {
							Console.WriteLine($"{extendTsvName},{relation.Key}:{key} not found in base {relation.Key}");
						}
					}
				}
			}
		}
	}
}
