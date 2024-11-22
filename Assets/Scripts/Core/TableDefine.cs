namespace Moh.Excel {
    /// <summary>
    /// 使用者種類
    /// </summary>
    public enum UserType {
        /// <summary>
        /// 禁用
        /// </summary>
        Forbid,

        /// <summary>
        /// client專用
        /// </summary>
        Client,

        /// <summary>
        /// server專用
        /// </summary>
        Server,

        /// <summary>
        /// client與server通用
        /// </summary>
        Both,
    }

    /// <summary>
    /// 資料表定義
    /// </summary>
    public class TableDefine {
        /// <summary>
        /// 欄位名稱列
        /// </summary>
        public static readonly int NAME_ROW = 1;

        /// <summary>
        /// 欄位型態列
        /// </summary>
        public static readonly int TYPE_ROW = 2;

        /// <summary>
        /// 用戶設定列
        /// </summary>
        public static readonly int USER_ROW = 3;

        /// <summary>
        /// 資料開始列
        /// </summary>
        public static readonly int DATA_START_ROW = 4;
    }
}