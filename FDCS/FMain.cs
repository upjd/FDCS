using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace FDCS
{
    public partial class FMain : Form
    {
        /// <summary>
        /// 表示是否弹出提示关闭当前应用程序
        /// 1表示不弹出.
        /// </summary>
        public static int Iclose;

        public FMain()
        {
            InitializeComponent();
        }

        private void FMain_Load(object sender, EventArgs e)
        {
            //显示主窗体
            var bwmMainChild = new FMainChild() { MdiParent = this };
            bwmMainChild.Show();
            //设置状态栏显示内容
            uStatusBar.Panels[4].MarqueeInfo.Start();
        }

        private void MdiManager_TabClosing(object sender, Infragistics.Win.UltraWinTabbedMdi.CancelableMdiTabEventArgs e)
        {
            //主界面不允许关闭
            if (e.Tab.Form.Text.Equals("Home"))
            {
                e.Cancel = true;
                return;
            }
            e.Cancel = MessageBox.Show(@"Confirm Close？", @"Yes/No", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes;
        }

        private void FMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Iclose == 1)
                return;
            e.Cancel = MessageBox.Show(@"Confirm Exit？", @"Yes/No", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes;
        }

        private void FMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void uExplorerBar_ItemClick(object sender, Infragistics.Win.UltraWinExplorerBar.ItemEventArgs e)
        {
            var cClass = e.Item.Tag.ToString();
            if (string.IsNullOrEmpty(cClass))
            {
                return;
            }
            var f = ExistForm(e.Item.Key);
            if (f == null) MenuDoubleClick(cClass);
            else f.Activate();
        }

        /// <summary>
        /// 通过点击的菜单来显示窗体
        /// </summary>
        /// <param name="cClass">str是表示当前点击的菜单对于的类名</param>
        public void MenuDoubleClick(string cClass)
        {
            var t = Type.GetType(cClass);
            if (t == null) return;
            try
            {
                var obj = Activator.CreateInstance(t);
                t.InvokeMember("MdiParent", BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty, null, obj, new object[] { this });
                t.InvokeMember("Show", BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod, null, obj, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw ex;
            }


        }

        /// <summary>
        /// 判断对应窗体是否已经打开
        /// </summary>
        /// <param name="str">传入当前查询的窗体的名称</param>
        /// <returns>返回已经存在的窗体</returns>
        private Form ExistForm(string str)
        {
            return MdiChildren.FirstOrDefault(f => f.Text == str);
        }

        private void tsbRelogin_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void tsbMenu_Click(object sender, EventArgs e)
        {
            uSplitterLeft.Collapsed = !uSplitterLeft.Collapsed;
        }

        private void 选项OToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
