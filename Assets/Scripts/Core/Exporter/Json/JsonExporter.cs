using System;
using System.Collections.Generic;
using LitJson;

namespace Moh.Excel.Exporter {
    /// <summary>
    /// 輸出json文件
    /// </summary>
    public class JsonExporter : DataExporter {
        /// <summary>
        /// 存檔副檔名
        /// </summary>
        protected override string ext { get { return ".json"; } }

        /// <summary>
        /// 目標用戶型態
        /// </summary>
        protected virtual UserType user { get { return UserType.Both; } }

        /// <summary>
        /// 存檔資料夾
        /// </summary>
        /// <remarks>名稱同用戶種類</remarks>
        protected override string folder { get { return Enum.GetName(user.GetType(), user) + "/"; } }

        /// <summary>
        /// 取得存檔內容
        /// </summary>
        /// <param name="table">解析過的資料表</param>
        /// <returns>是否成功, 存檔內容</returns>
        protected override (bool, string) GetSaveData(ParseTableData table) {
            var res = new Dictionary<string, object>();

            // 資料總數
            var count = table.rows - TableDefine.DATA_START_ROW;

            // 資料內容
            for (int i = 0; i < count; i++) {
                var data = GetData(table, i);

                // 無此資料
                if (data == null || data.Count <= 0) {
                    continue;
                }

                // 一定要有id欄位
                if (data.TryGetValue("id", out var id) == false) {
                    throw new Exception(string.Format("export json {0} failed, id not found", table.name));
                }

                // 此筆資料有無開放
                if (data.TryGetValue("open", out var open)) {
                    var str = open.ToString();

                    // 值為數字型態
                    if (int.TryParse(str, out var openInt) && openInt == 0) {
                        continue;
                    }
                    // 值為字串型
                    else if (bool.TryParse(str, out var openBool) && openBool == false) {
                        continue;
                    }
                }

                res.Add(id.ToString(), data);
            }

            if (res.Count <= 0) {
                return (false, string.Empty);
            }

            return (true, JsonMapper.ToJson(res));
        }

        /// <summary>
        /// 取得資料
        /// </summary>
        /// <param name="table">解析過的資料表</param>
        /// <param name="idx">資料索引</param>
        private Dictionary<string, object> GetData(ParseTableData table, int idx) {
            var res = new Dictionary<string, object>();

            // 將各欄的資料組裝成此筆資料回傳
            for (var i = 0; i < table.columns; i++) {
                var col = table.contain[i];
                var data = col.contain[idx];

                // 處理目標用戶
                if (col.user != UserType.Both && col.user != user) {
                    continue;
                }

                // 欄位名稱
                var name = col.name;

                // 此欄位為編號
                if (name.ToLower() == "id") {
                    name = "id";  // 強制改名
                }

                // 此欄位為開放
                if (name.ToLower() == "open") {
                    name = "open";  // 強制改名
                }

                // 陣列處理
                if (col.isAyBegin) {
                    var (end, list) = GetDataAy(table, idx, i);
                    res.Add(name, list);

                    // 迴圈跳至陣列結尾
                    i = end;

                    continue;
                }

                // 欄位名稱重複
                if (res.ContainsKey(name)) {
                    throw new Exception(string.Format("export json {0} failed, field {1} is repeat", table.name, name));
                }

                res.Add(name, data);
            }

            return res;
        }

        /// <summary>
        /// 取得陣列資料
        /// </summary>
        /// <param name="table">解析過的資料表</param>
        /// <param name="idx">資料索引</param>
        /// <returns>陣列結尾欄, 陣列資料</returns>
        private (int, List<object>) GetDataAy(ParseTableData table, int idx, int start) {
            var res = new List<object>();

            for (int i = start; i < table.columns; i++) {
                var col = table.contain[i];
                res.Add(col.contain[idx]);

                if (col.isAyEnd) {
                    return (i, res);
                }
            }

            // 找不到陣列結尾
            throw new Exception(string.Format("export json {0} failed, field {1} array end not found", table.name, table.contain[start].name));
        }
    }
}
