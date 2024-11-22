using Moh.Excel.Exporter;
using UnityEngine;

namespace Moh.Excel.Test {
    /// <summary>
    /// 測試入口
    /// </summary>
    public class TestEntry : MonoBehaviour {
        /// <summary>
        /// 
        /// </summary>
        private void Start() {
            var table = new ParseTableData("Assets/Excel/Input/", "TestExcel.xlsx");
            table.Export("Assets/Excel/Output/", new JsonExporterC(), new JsonExporterS());
        }
    }
}
