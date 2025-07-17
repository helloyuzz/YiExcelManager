using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

public partial class ExcelSheetNavigator : UserControl {
    private Excel.Application excel;
    private ListBox sheetListBox;
    private LinkLabel linkLabel;

    public ExcelSheetNavigator(Excel.Application app) {
        excel = app;
        InitializeComponent();
        InitializeSheetList();
    }

    private void InitializeComponent() {
        sheetListBox = new ListBox();
        sheetListBox.Dock = DockStyle.Fill;
        sheetListBox.Font = new System.Drawing.Font("Microsoft YaHei UI", 10);
        sheetListBox.SelectedIndexChanged += SheetListBox_SelectedIndexChanged;
        sheetListBox.ItemHeight = 23;
        this.Controls.Add(sheetListBox);

        linkLabel= new LinkLabel();
        linkLabel.Text = "Author:Yulinyi";
        linkLabel.Dock = DockStyle.Bottom;
        this.Controls.Add(linkLabel);
    }

    private void InitializeSheetList() {
        sheetListBox.Items.Clear();
        if (excel.ActiveWorkbook == null) {
            MessageBox.Show("请先打开一个工作簿。", "YiExcelManager提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        foreach (Excel.Worksheet sheet in excel.ActiveWorkbook.Worksheets) {
            sheetListBox.Items.Add(sheet.Name);
        }
    }

    private void SheetListBox_SelectedIndexChanged(object sender, EventArgs e) {
        string selectedSheet = sheetListBox.SelectedItem as string;
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
}
