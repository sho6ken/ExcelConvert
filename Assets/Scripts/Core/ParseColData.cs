using System;
using System.Collections.Generic;

namespace Moh.Excel {
    /// <summary>
    /// 解析整欄數據
    /// </summary>
    public class ParseColData {
        /// <summary>
        /// 欄位名稱
        /// </summary>
        public string name { get; private set; }

        /// <summary>
        /// 資料型態
        /// </summary>
        public Type type { get; private set; }

        /// <summary>
        /// 是否為陣列型資料開頭
        /// </summary>
        public bool isAyBegin { get; private set; }

        /// <summary>
        /// 是否為陣列型資料結尾
        /// </summary>
        public bool isAyEnd { get; private set; }

        /// <summary>
        /// 用戶種類
        /// </summary>
        public UserType user { get; private set; }

        /// <summary>
        /// 此欄位是否開放
        /// </summary>
        public bool opened { get { return user != UserType.Forbid; } }

        /// <summary>
        /// 實際數據資料
        /// </summary>
        public List<object> contain { get; private set; }

        /// <summary>
        /// 資料總筆數
        /// </summary>
        public int count { get { return contain.Count; } }

        /// <summary>
        /// 
        /// </summary>
        public ParseColData(List<object> data) {
            ParseName(data);
            ParseType(data);
            ParseUser(data);
            ParseContain(data);
        }

        /// <summary>
        /// 解析欄位名稱
        /// </summary>
        private void ParseName(List<object> data) {
            name = data[TableDefine.NAME_ROW].ToString();
        }

        /// <summary>
        /// 解析數據型態
        /// </summary>
        private void ParseType(List<object> data) {
            var str = data[TableDefine.TYPE_ROW].ToString().ToLower();

            isAyBegin = false;
            isAyEnd = false;

            // 陣列開頭
            if (str.StartsWith("[")) {
                isAyBegin = true;
                str.Substring(1);
            }

            // 陣列結尾
            if (str.EndsWith("]")) {
                isAyEnd = true;
                str.Substring(0, str.Length - 1);
            }

            // 型態判斷
            switch (str) {
                case "byte":
                case "short":
                case "int":
                    type = typeof(int);
                    break;

                case "bool":
                    type = typeof(bool);
                    break;

                case "string":
                    type = typeof(string);
                    break;

                // 當型態錯誤時禁止使用
                default:
                    type = null;
                    user = UserType.Forbid;
                    break;
            }
        }

        /// <summary>
        /// 解析用戶設定
        /// </summary>
        private void ParseUser(List<object> data) {
            var str = data[TableDefine.USER_ROW].ToString().ToLower();

            switch (str) {
                case "c" : user = UserType.Client; break;
                case "s" : user = UserType.Server; break;
                case "cs": user = UserType.Both;   break;
                case "sc": user = UserType.Both;   break;
                default  : user = UserType.Forbid; break;
            }
        }

        /// <summary>
        /// 解析實際數據資料
        /// </summary>
        private void ParseContain(List<object> data) {
            var count = data.Count;

            // 無資料
            if (count < TableDefine.DATA_START_ROW) {
                throw new Exception(string.Format("parse col data {0} failed, data not found", name));
            }

            contain = new List<object>();
            
            for (var i = TableDefine.DATA_START_ROW; i < count; i++) {
                contain.Add(data[i].ToString());
            }
        }
    }
}
