using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YiExcelManager {
    public partial class ExcelSheetPanel : UserControl {
        private Excel.Application excel;
        public ExcelSheetPanel(Excel.Application app) {
            InitializeComponent();
            excel = app;
        }



        private void InitializeSheetList() {
            listBox1.Items.Clear();
            if (excel.ActiveWorkbook == null) {
                MessageBox.Show("请先打开一个工作簿。", "YiExcelManager提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            foreach (Excel.Worksheet sheet in excel.ActiveWorkbook.Worksheets) {
                listBox1.Items.Add(sheet.Name);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e) {
            string selectedSheet = listBox1.SelectedItem as string;
            if(excel.ActiveWorkbook == null) {
                MessageBox.Show("请先打开一个工作簿。", "YiExcelManager提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!string.IsNullOrEmpty(selectedSheet)) {
                foreach (Excel.Worksheet sheet in excel.ActiveWorkbook.Worksheets) {
                    if (sheet.Name == selectedSheet) {
                        sheet.Activate();
                        break;
                    }
                }
            }
        }

        // Call this method if you want to refresh the sheet list (e.g., after adding/removing sheets)
        public void RefreshSheets() {
            InitializeSheetList();
        }
        public void CloseSheets() {
            listBox1.Items.Clear();
        }
    }
}
