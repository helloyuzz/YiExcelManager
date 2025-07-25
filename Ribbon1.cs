﻿using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YiExcelManager {
    public partial class Ribbon1 {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {

        }

        private void btn_About_Click(object sender, RibbonControlEventArgs e) {
            AboutBox aboutBox = new AboutBox();
            aboutBox.ShowDialog();
        }

        private void btn_Display_Click(object sender, RibbonControlEventArgs e) {
            foreach (var panel in Globals.ThisAddIn.CustomTaskPanes) {
                if (panel.Title == "YiExcelManager") {
                    panel.Visible = !panel.Visible;
                }
            }
        }
    }
}
