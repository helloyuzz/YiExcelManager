using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System;

namespace YiExcelManager {
    public partial class ThisAddIn {
        private CustomTaskPane sheetNavigationPane;
        private ExcelSheetPanel navigator;
        private Microsoft.Office.Interop.Excel.Workbook currentWorkbook;


        private void ThisAddIn_Startup(object sender, EventArgs e) {
            navigator = new ExcelSheetPanel(this.Application);
            sheetNavigationPane = this.CustomTaskPanes.Add(navigator, "YiExcelManager");
            sheetNavigationPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            sheetNavigationPane.Width = 468;
            sheetNavigationPane.Visible = true;

            // 订阅事件
            this.Application.SheetActivate += Application_SheetActivate;

            // 订阅当前活动工作簿的事件
            currentWorkbook = this.Application.ActiveWorkbook;
            if (currentWorkbook != null) {
                currentWorkbook.NewSheet += Workbook_NewSheet;
                currentWorkbook.SheetActivate += Workbook_SheetActivate;
            }
            this.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_WorkbookActivate(Workbook Wb) {
            navigator.RefreshSheets();
        }

        private void Workbook_NewSheet(object sh) {
            navigator.RefreshSheets();
        }

        private void Workbook_SheetActivate(object sh) {
            navigator.RefreshSheets();
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
            // 清理事件监听
            if (currentWorkbook != null) {
                currentWorkbook.NewSheet -= Workbook_NewSheet;
                currentWorkbook.SheetActivate -= Workbook_SheetActivate;
            }

            this.Application.SheetActivate -= Application_SheetActivate;
        }


        // 激活Sheet时刷新（可用于检测重命名）
        private void Application_SheetActivate(object sh) {
            navigator.RefreshSheets();
        }

        #region VSTO 生成的代码
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}