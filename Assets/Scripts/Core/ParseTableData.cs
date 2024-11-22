using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel;

namespace Moh.Excel {
    /// <summary>
    /// 解析excel資料表
    /// </summary>
    public class ParseTableData {
        /// <summary>
        /// 表名
        /// </summary>
        public string name { get; private set; }

        /// <summary>
        /// 欄數
        /// </summary>
        public int columns { get; private set; }

        /// <summary>
        /// 列數
        /// </summary>
        public int rows { get; private set; }

        /// <summary>
        /// 各欄資料
        /// </summary>
        public List<ParseColData> contain { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">完整路徑(含檔名)</param>
        public ParseTableData(string path) {
            if (path.EndsWith(".xlsx") == false) {
                throw new Exception(string.Format("read table {0} failed, ext not found", path));
            }

            var idx = path.LastIndexOf("/");
            
            if (idx == -1) {
                throw new Exception(string.Format("read table {0} failed, path is illegal", path));
            }

            // 表名
            name = path.Substring(idx + 1, path.Length - idx - 1);
            name = name.Replace(".xlsx", string.Empty);

            // 初始化
            Init(path);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">檔案路徑</param>
        /// <param name="name">檔名</param>
        /// <remarks>只會處理第一個切頁</remarks>
        public ParseTableData(string path, string name) {
            if (path.EndsWith("/") == false) {
                path += "/";
            }

            if (name.EndsWith(".xlsx") == false) {
                name += ".xlsx";
            }

            // 表名
            this.name = name.Replace(".xlsx", string.Empty);

            // 初始化
            Init(path + name);
        }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="path">完整路徑(含檔名)</param>
        private void Init(string path) {
            // 讀檔
            var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            var data = reader.AsDataSet();

            if (data.Tables.Count <= 0) {
                throw new Exception(string.Format("read table {0} failed, sheet is null", path));
            }

            // 解析資料表
            ParseTable(data.Tables[0]);
        }

        /// <summary>
        /// 解析資料表
        /// </summary>
        private void ParseTable(DataTable table) {
            if (table == null) {
                throw new Exception(string.Format("parse table {0} failed, table is null", name));
            }

            contain = new List<ParseColData>();

            // 計算內容欄數
            for (var i = 0; i < table.Columns.Count; i++) {
                var str = table.Rows[0][i].ToString().ToLower();

                if (str == "eof" || string.IsNullOrEmpty(str)) {
                    columns = i;
                    break;
                }
            }

            // 計算內容列數
            for (var i = 0; i < table.Rows.Count; i++) {
                var str = table.Rows[i][0].ToString().ToLower();

                if (str == "eof" || string.IsNullOrEmpty(str)) {
                    rows = i;
                    break;
                }
            }

            // 無資料
            if (rows < TableDefine.DATA_START_ROW) {
                throw new Exception(string.Format("parse table {0} failed, data not found", name));
            }

            // 解析各欄
            for (var i = 0; i < columns; i++) {
                var data = GetColData(table, i);
                contain.Add(new ParseColData(data));
            }
        }

        /// <summary>
        /// 取得此欄所有資料
        /// </summary>
        private List<object> GetColData(DataTable table, int col) {
            var data = new List<object>();

            for (var i = 0; i < rows; i++) {
                data.Add(table.Rows[i][col]);
            }

            return data;
        }

        /// <summary>
        /// 輸出文件
        /// </summary>
        /// <param name="path">輸出路徑</param>
        /// <param name="exporters">各資料輸出器</param>
        public void Export(string path, params DataExporter[] exporters) {
            foreach (var elm in exporters) {
                elm.Execute(path, this);
            }
        }
    }
}
