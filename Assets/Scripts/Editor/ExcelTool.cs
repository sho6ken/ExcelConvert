using System.Collections.Generic;
using Moh.Excel.Exporter;
using UnityEditor;
using UnityEngine;

namespace Moh.Excel.Tool {
    /// <summary>
    /// excel轉檔工具
    /// </summary>
    public class ExcelTool : EditorWindow {
        /// <summary>
        /// 單例
        /// </summary>
        private static ExcelTool _inst = null;

        /// <summary>
        /// excel文件列表
        /// </summary>
        private static List<string> _files = new List<string>();

        /// <summary>
        /// 檔案總數
        /// </summary>
        private static int _count { get { return _files.Count; } }

        /// <summary>
        /// 輸出路徑
        /// </summary>
        private static string _dest = string.Empty;

        /// <summary>
        /// 顯示視窗
        /// </summary>
        [MenuItem("Assets/墨/表格轉檔")]
        private static void ShowForm() {
            _inst = GetWindow<ExcelTool>();
            _dest = Application.dataPath;

            Load();

            _inst.Show();
        }

        /// <summary>
        /// 讀取文件
        /// </summary>
        private static void Load() {
            _files.Clear();

            var selects = (object[])Selection.objects;

            // 沒有選到的文件
            if (selects.Length <= 0) {
                return;
            }

            foreach (var elm in selects) {
                var path = AssetDatabase.GetAssetPath((Object)elm);

                if (path.EndsWith(".xlsx")) {
                    _files.Add(path);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void OnGUI() {
            // 顯示輸出路徑
            EditorGUILayout.LabelField(string.Format("export folder: {0}", _dest));

            // 更換輸出路徑
            if (GUILayout.Button("change export folder")) {
                _dest = EditorUtility.OpenFolderPanel("select export folder", Application.dataPath, "");
            }

            // 繪製文件列表
            DrawFiles();

            // 執行
            if (GUILayout.Button("execute")) {
                Execute();
            }
        }

        /// <summary>
        /// 繪製文件列表
        /// </summary>
        private void DrawFiles() {
            if (_files.Count <= 0) {
                EditorGUILayout.LabelField("no excel files selected");
                return;
            }

            EditorGUILayout.LabelField(string.Format("total {0} excel files selected:", _count));

            GUILayout.BeginVertical();
            GUILayout.BeginScrollView(Vector2.zero, false, true, GUILayout.Height(250));

            // 繪製文件列表
            for (var i = 0; i < _count; i++) {
                GUILayout.BeginHorizontal();
                GUILayout.Label(string.Format("{0}: {1}", i, _files[i]));
                GUILayout.EndHorizontal();
            }

            GUILayout.EndScrollView();
            GUILayout.EndVertical();
        }

        /// <summary>
        /// 執行轉檔
        /// </summary>
        private void Execute() {
            // 未設定輸出路徑
            if (string.IsNullOrEmpty(_dest)) {
                Debug.LogError("execute export excel failed, dest is null");
                return;
            }

            // 清除log
            Debug.ClearDeveloperConsole();

            Debug.LogFormat("start export all excel, total {0} excel files", _count);

            for (int i = 0; i < _count; i++) {
                Debug.LogFormat("{0}: {1}", i, _files[i]);

                // 輸出文件
                var table = new ParseTableData(_files[i]);
                table.Export(_dest, new JsonExporterC(), new JsonExporterS());

                // 刷新資源顯示
                AssetDatabase.Refresh();
            }

            Debug.Log("export all excel done");

            _inst.Close();
        }

        /// <summary>
        /// 當選取內容發生變化時
        /// </summary>
        private void OnSelectionChange() {
            Show();
            Load();
            Repaint();
        }
    }
}
