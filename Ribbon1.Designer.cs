namespace YiExcelManager {
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btn_Display = this.Factory.CreateRibbonButton();
            this.btn_About = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_Display);
            this.group1.Items.Add(this.btn_About);
            this.group1.Label = "YiExcelManager";
            this.group1.Name = "group1";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "bookmark_all.png");
            this.imageList1.Images.SetKeyName(1, "create_bookmark.png");
            this.imageList1.Images.SetKeyName(2, "new_bookmark.png");
            // 
            // btn_Display
            // 
            this.btn_Display.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Display.Image = global::YiExcelManager.Properties.Resources.nav_menu;
            this.btn_Display.Label = "显示导航";
            this.btn_Display.Name = "btn_Display";
            this.btn_Display.ScreenTip = "显示导航";
            this.btn_Display.ShowImage = true;
            // 
            // btn_About
            // 
            this.btn_About.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_About.Image = global::YiExcelManager.Properties.Resources.about;
            this.btn_About.Label = "关于";
            this.btn_About.Name = "btn_About";
            this.btn_About.ScreenTip = "关于";
            this.btn_About.ShowImage = true;
            this.btn_About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_About_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Display;
        private System.Windows.Forms.ImageList imageList1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_About;
    }

    partial class ThisRibbonCollection {
        internal Ribbon1 Ribbon1 {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
