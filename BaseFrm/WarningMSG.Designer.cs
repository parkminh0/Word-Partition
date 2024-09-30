namespace WordPartition
{
    partial class WarningMSG
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WarningMSG));
            this.mmeMsg = new DevExpress.XtraEditors.MemoEdit();
            this.btnClose = new DevExpress.XtraEditors.SimpleButton();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            ((System.ComponentModel.ISupportInitialize)(this.mmeMsg.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // mmeMsg
            // 
            this.mmeMsg.Location = new System.Drawing.Point(5, 5);
            this.mmeMsg.Name = "mmeMsg";
            this.mmeMsg.Properties.ReadOnly = true;
            this.mmeMsg.Size = new System.Drawing.Size(437, 185);
            this.mmeMsg.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.Location = new System.Drawing.Point(196, 196);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 30);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "창닫기";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // panelControl1
            // 
            this.panelControl1.Controls.Add(this.mmeMsg);
            this.panelControl1.Controls.Add(this.btnClose);
            this.panelControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelControl1.Location = new System.Drawing.Point(0, 0);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(447, 231);
            this.panelControl1.TabIndex = 2;
            // 
            // WarningMSG
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(447, 231);
            this.Controls.Add(this.panelControl1);
            this.Name = "WarningMSG";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "WarningMSG";
            ((System.ComponentModel.ISupportInitialize)(this.mmeMsg.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.MemoEdit mmeMsg;
        private DevExpress.XtraEditors.SimpleButton btnClose;
        private DevExpress.XtraEditors.PanelControl panelControl1;
    }
}