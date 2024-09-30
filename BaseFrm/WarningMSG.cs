using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace WordPartition
{
    public partial class WarningMSG : DevExpress.XtraEditors.XtraForm
    {
        public WarningMSG()
        {
            InitializeComponent();
        }

        delegate void ShowMsg(string m);

        public void MSG(string msg)
        {

            mmeMsg.Text = msg;
            this.ShowDialog(Program.mainApp);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}