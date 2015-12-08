namespace FDCS
{
    partial class FMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FMain));
            Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup ultraExplorerBarGroup1 = new Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup();
            Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem ultraExplorerBarItem14 = new Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem();
            Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
            Infragistics.Win.UltraWinStatusBar.UltraStatusPanel ultraStatusPanel6 = new Infragistics.Win.UltraWinStatusBar.UltraStatusPanel();
            Infragistics.Win.UltraWinStatusBar.UltraStatusPanel ultraStatusPanel7 = new Infragistics.Win.UltraWinStatusBar.UltraStatusPanel();
            Infragistics.Win.UltraWinStatusBar.UltraStatusPanel ultraStatusPanel8 = new Infragistics.Win.UltraWinStatusBar.UltraStatusPanel();
            Infragistics.Win.UltraWinStatusBar.UltraStatusPanel ultraStatusPanel9 = new Infragistics.Win.UltraWinStatusBar.UltraStatusPanel();
            Infragistics.Win.UltraWinStatusBar.UltraStatusPanel ultraStatusPanel10 = new Infragistics.Win.UltraWinStatusBar.UltraStatusPanel();
            Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
            this.tsMain = new System.Windows.Forms.ToolStrip();
            this.tsbRelogin = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbMenu = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbConnect = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.msMain = new System.Windows.Forms.MenuStrip();
            this.文件FToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmRelogin = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.工具TToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiShowCalculater = new System.Windows.Forms.ToolStripMenuItem();
            this.选项OToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.帮助HToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.内容CToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.关于AToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panelLeft = new System.Windows.Forms.Panel();
            this.uExplorerBar = new Infragistics.Win.UltraWinExplorerBar.UltraExplorerBar();
            this.uSplitterLeft = new Infragistics.Win.Misc.UltraSplitter();
            this.uStatusBar = new Infragistics.Win.UltraWinStatusBar.UltraStatusBar();
            this.MdiManager = new Infragistics.Win.UltraWinTabbedMdi.UltraTabbedMdiManager(this.components);
            this.tsMain.SuspendLayout();
            this.msMain.SuspendLayout();
            this.panelLeft.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.uExplorerBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uStatusBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.MdiManager)).BeginInit();
            this.SuspendLayout();
            // 
            // tsMain
            // 
            this.tsMain.AutoSize = false;
            this.tsMain.BackColor = System.Drawing.SystemColors.Control;
            this.tsMain.BackgroundImage = global::FDCS.Properties.Resources.toolbarBk;
            this.tsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbRelogin,
            this.toolStripSeparator1,
            this.tsbMenu,
            this.toolStripSeparator2,
            this.tsbConnect,
            this.toolStripSeparator6});
            this.tsMain.Location = new System.Drawing.Point(0, 25);
            this.tsMain.Name = "tsMain";
            this.tsMain.Size = new System.Drawing.Size(984, 25);
            this.tsMain.TabIndex = 12;
            this.tsMain.Text = "toolStrip1";
            // 
            // tsbRelogin
            // 
            this.tsbRelogin.AutoSize = false;
            this.tsbRelogin.Image = ((System.Drawing.Image)(resources.GetObject("tsbRelogin.Image")));
            this.tsbRelogin.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbRelogin.Name = "tsbRelogin";
            this.tsbRelogin.Size = new System.Drawing.Size(80, 22);
            this.tsbRelogin.Text = "Exit";
            this.tsbRelogin.ToolTipText = "Exit";
            this.tsbRelogin.Click += new System.EventHandler(this.tsbRelogin_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // tsbMenu
            // 
            this.tsbMenu.AutoSize = false;
            this.tsbMenu.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbMenu.Image = global::FDCS.Properties.Resources.previous_page;
            this.tsbMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbMenu.Name = "tsbMenu";
            this.tsbMenu.Size = new System.Drawing.Size(80, 22);
            this.tsbMenu.Text = "Navigation";
            this.tsbMenu.Click += new System.EventHandler(this.tsbMenu_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // tsbConnect
            // 
            this.tsbConnect.AutoSize = false;
            this.tsbConnect.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbConnect.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbConnect.Name = "tsbConnect";
            this.tsbConnect.Size = new System.Drawing.Size(80, 22);
            this.tsbConnect.Text = "Contact";
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            this.toolStripSeparator6.Size = new System.Drawing.Size(6, 25);
            // 
            // msMain
            // 
            this.msMain.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(194)))), ((int)(((byte)(216)))), ((int)(((byte)(240)))));
            this.msMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件FToolStripMenuItem,
            this.工具TToolStripMenuItem,
            this.帮助HToolStripMenuItem});
            this.msMain.Location = new System.Drawing.Point(0, 0);
            this.msMain.Name = "msMain";
            this.msMain.Size = new System.Drawing.Size(984, 25);
            this.msMain.TabIndex = 11;
            this.msMain.Text = "menuStrip1";
            // 
            // 文件FToolStripMenuItem
            // 
            this.文件FToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmRelogin,
            this.toolStripSeparator});
            this.文件FToolStripMenuItem.Name = "文件FToolStripMenuItem";
            this.文件FToolStripMenuItem.Size = new System.Drawing.Size(76, 21);
            this.文件FToolStripMenuItem.Text = "System(&S)";
            // 
            // tsmRelogin
            // 
            this.tsmRelogin.Name = "tsmRelogin";
            this.tsmRelogin.Size = new System.Drawing.Size(111, 22);
            this.tsmRelogin.Text = "Exit(&E)";
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(108, 6);
            // 
            // 工具TToolStripMenuItem
            // 
            this.工具TToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiShowCalculater,
            this.选项OToolStripMenuItem});
            this.工具TToolStripMenuItem.Name = "工具TToolStripMenuItem";
            this.工具TToolStripMenuItem.Size = new System.Drawing.Size(61, 21);
            this.工具TToolStripMenuItem.Text = "Tool(&T)";
            // 
            // tsmiShowCalculater
            // 
            this.tsmiShowCalculater.Name = "tsmiShowCalculater";
            this.tsmiShowCalculater.Size = new System.Drawing.Size(150, 22);
            this.tsmiShowCalculater.Text = "Calculator(&C)";
            // 
            // 选项OToolStripMenuItem
            // 
            this.选项OToolStripMenuItem.Name = "选项OToolStripMenuItem";
            this.选项OToolStripMenuItem.Size = new System.Drawing.Size(150, 22);
            this.选项OToolStripMenuItem.Text = "Option(&O)";
            // 
            // 帮助HToolStripMenuItem
            // 
            this.帮助HToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.内容CToolStripMenuItem,
            this.toolStripSeparator5,
            this.关于AToolStripMenuItem});
            this.帮助HToolStripMenuItem.Name = "帮助HToolStripMenuItem";
            this.帮助HToolStripMenuItem.Size = new System.Drawing.Size(64, 21);
            this.帮助HToolStripMenuItem.Text = "Help(&H)";
            // 
            // 内容CToolStripMenuItem
            // 
            this.内容CToolStripMenuItem.Name = "内容CToolStripMenuItem";
            this.内容CToolStripMenuItem.Size = new System.Drawing.Size(149, 22);
            this.内容CToolStripMenuItem.Text = "内容(&C)";
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(146, 6);
            // 
            // 关于AToolStripMenuItem
            // 
            this.关于AToolStripMenuItem.Name = "关于AToolStripMenuItem";
            this.关于AToolStripMenuItem.Size = new System.Drawing.Size(149, 22);
            this.关于AToolStripMenuItem.Text = "关于我们(&A)...";
            // 
            // panelLeft
            // 
            this.panelLeft.BackColor = System.Drawing.Color.White;
            this.panelLeft.Controls.Add(this.uExplorerBar);
            this.panelLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelLeft.Location = new System.Drawing.Point(0, 50);
            this.panelLeft.Name = "panelLeft";
            this.panelLeft.Size = new System.Drawing.Size(222, 489);
            this.panelLeft.TabIndex = 13;
            // 
            // uExplorerBar
            // 
            this.uExplorerBar.Dock = System.Windows.Forms.DockStyle.Fill;
            ultraExplorerBarItem14.Key = "Data Transfer";
            appearance1.Image = global::FDCS.Properties.Resources.chart_8;
            ultraExplorerBarItem14.Settings.AppearancesLarge.Appearance = appearance1;
            ultraExplorerBarItem14.Tag = "FDCS.Work01TransferVouch";
            ultraExplorerBarItem14.Text = "Data Transfer";
            ultraExplorerBarGroup1.Items.AddRange(new Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem[] {
            ultraExplorerBarItem14});
            ultraExplorerBarGroup1.Key = "Data Central";
            ultraExplorerBarGroup1.Text = "Data Central";
            this.uExplorerBar.Groups.AddRange(new Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup[] {
            ultraExplorerBarGroup1});
            this.uExplorerBar.GroupSettings.Style = Infragistics.Win.UltraWinExplorerBar.GroupStyle.LargeImagesWithText;
            this.uExplorerBar.Location = new System.Drawing.Point(0, 0);
            this.uExplorerBar.Name = "uExplorerBar";
            this.uExplorerBar.ShowDefaultContextMenu = false;
            this.uExplorerBar.Size = new System.Drawing.Size(222, 489);
            this.uExplorerBar.TabIndex = 2;
            this.uExplorerBar.ViewStyle = Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarViewStyle.Office2007;
            this.uExplorerBar.ItemClick += new Infragistics.Win.UltraWinExplorerBar.ItemClickEventHandler(this.uExplorerBar_ItemClick);
            // 
            // uSplitterLeft
            // 
            this.uSplitterLeft.Location = new System.Drawing.Point(222, 50);
            this.uSplitterLeft.Name = "uSplitterLeft";
            this.uSplitterLeft.RestoreExtent = 191;
            this.uSplitterLeft.Size = new System.Drawing.Size(10, 489);
            this.uSplitterLeft.TabIndex = 14;
            // 
            // uStatusBar
            // 
            this.uStatusBar.Location = new System.Drawing.Point(0, 539);
            this.uStatusBar.Name = "uStatusBar";
            ultraStatusPanel6.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            ultraStatusPanel6.DateTimeFormat = "yyyy-MM-dd hh:mm:ss";
            ultraStatusPanel6.Style = Infragistics.Win.UltraWinStatusBar.PanelStyle.Date;
            ultraStatusPanel6.Width = 150;
            ultraStatusPanel7.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            ultraStatusPanel7.Key = "tssl_Lname";
            ultraStatusPanel7.Text = "GE Digital";
            ultraStatusPanel7.Width = 150;
            ultraStatusPanel8.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            ultraStatusPanel8.Key = "tssl_Lserver";
            ultraStatusPanel8.Text = "China";
            ultraStatusPanel8.Width = 300;
            ultraStatusPanel9.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
            ultraStatusPanel9.Key = "tssbtnShow";
            ultraStatusPanel9.Style = Infragistics.Win.UltraWinStatusBar.PanelStyle.Button;
            ultraStatusPanel9.Text = "Version:1.0.0.0";
            ultraStatusPanel9.Width = 120;
            appearance3.TextHAlignAsString = "Right";
            ultraStatusPanel10.Appearance = appearance3;
            ultraStatusPanel10.Key = "cCompany";
            ultraStatusPanel10.MarqueeInfo.Delay = 50;
            ultraStatusPanel10.SizingMode = Infragistics.Win.UltraWinStatusBar.PanelSizingMode.Spring;
            ultraStatusPanel10.Style = Infragistics.Win.UltraWinStatusBar.PanelStyle.Marquee;
            ultraStatusPanel10.Text = "GE | Imagination at Work ";
            ultraStatusPanel10.Width = 300;
            ultraStatusPanel10.WrapText = Infragistics.Win.DefaultableBoolean.False;
            this.uStatusBar.Panels.AddRange(new Infragistics.Win.UltraWinStatusBar.UltraStatusPanel[] {
            ultraStatusPanel6,
            ultraStatusPanel7,
            ultraStatusPanel8,
            ultraStatusPanel9,
            ultraStatusPanel10});
            this.uStatusBar.Size = new System.Drawing.Size(984, 23);
            this.uStatusBar.TabIndex = 15;
            this.uStatusBar.ViewStyle = Infragistics.Win.UltraWinStatusBar.ViewStyle.Office2007;
            // 
            // MdiManager
            // 
            this.MdiManager.AllowHorizontalTabGroups = false;
            this.MdiManager.AllowVerticalTabGroups = false;
            this.MdiManager.MdiParent = this;
            this.MdiManager.ViewStyle = Infragistics.Win.UltraWinTabbedMdi.ViewStyle.Office2007;
            this.MdiManager.TabClosing += new Infragistics.Win.UltraWinTabbedMdi.CancelableMdiTabEventHandler(this.MdiManager_TabClosing);
            // 
            // FMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 562);
            this.Controls.Add(this.uSplitterLeft);
            this.Controls.Add(this.panelLeft);
            this.Controls.Add(this.tsMain);
            this.Controls.Add(this.msMain);
            this.Controls.Add(this.uStatusBar);
            this.Icon = global::FDCS.Properties.Resources.Mine;
            this.IsMdiContainer = true;
            this.Name = "FMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FDCS Home Page";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FMain_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FMain_FormClosed);
            this.Load += new System.EventHandler(this.FMain_Load);
            this.tsMain.ResumeLayout(false);
            this.tsMain.PerformLayout();
            this.msMain.ResumeLayout(false);
            this.msMain.PerformLayout();
            this.panelLeft.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.uExplorerBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uStatusBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.MdiManager)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip tsMain;
        private System.Windows.Forms.ToolStripButton tsbRelogin;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton tsbMenu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton tsbConnect;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.MenuStrip msMain;
        private System.Windows.Forms.ToolStripMenuItem 文件FToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmRelogin;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripMenuItem 工具TToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiShowCalculater;
        private System.Windows.Forms.ToolStripMenuItem 选项OToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 帮助HToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 内容CToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripMenuItem 关于AToolStripMenuItem;
        private System.Windows.Forms.Panel panelLeft;
        private Infragistics.Win.UltraWinExplorerBar.UltraExplorerBar uExplorerBar;
        private Infragistics.Win.Misc.UltraSplitter uSplitterLeft;
        private Infragistics.Win.UltraWinStatusBar.UltraStatusBar uStatusBar;
        private Infragistics.Win.UltraWinTabbedMdi.UltraTabbedMdiManager MdiManager;
    }
}