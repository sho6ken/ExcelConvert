using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Moh.Excel {
    /// <summary>
    /// 數據輸出器
    /// </summary>
    public abstract class DataExporter {
        /// <summary>
        /// 存檔資料夾
        /// </summary>
        protected virtual string folder { get { return string.Empty; } }

        /// <summary>
        /// 存檔副檔名
        /// </summary>
        protected abstract string ext { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="table">輸出路徑</param>
        /// <param name="table">解析過的資料表</param>
        public virtual void Execute(string path, ParseTableData table) {
            if (table == null) {
                throw new Exception("export table failed, table is null");
            }

            if (path.EndsWith("/") == false) {
                path += "/";
            }

            // 含指定資料夾
            path += folder;

            // 創建指定資料夾
            if (Directory.Exists(path) == false) {
                Directory.CreateDirectory(path);
            }

            // 含檔名
            path += table.name + ext;

            // 取得存檔內容
            var (res, data) = GetSaveData(table);

            if (res == false) {
                throw new Exception(string.Format("export table {0} failed, data is null", path));
            }

            // 解決中文字變亂碼的問題
            data = new Regex(@"(?i)\\[uU]([0-9a-f]{4})").Replace(data, delegate (Match match) { 
                return ((char)Convert.ToInt32(match.Groups[1].Value, 16)).ToString();
            });

            // 存檔
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write)) {
                using (var writer = new StreamWriter(stream, Encoding.Unicode)) {
                    writer.Write(data);
                }
            }
        }

        /// <summary>
        /// 取得存檔內容
        /// </summary>
        /// <param name="table">解析過的資料表</param>
        /// <returns>是否成功, 存檔內容</returns>
        protected abstract (bool, string) GetSaveData(ParseTableData table);
    }
}
