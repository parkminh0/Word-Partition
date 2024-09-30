using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraReports.UI;

namespace WordPartition
{
    public partial class BaseForm : XtraForm
    {
        //public BaseForm()
        //{
        //    InitializeComponent();
        //}
        public OptionInfo OCF
        {
            get
            {
                return Program.Option;
            }
        }

        public WarningMSG WMSG
        {
            get
            {
                return Program.WMSG;
            }
        }

        private void BaseForm_Load(object sender, EventArgs e)
        {

        }

        static object lastEdit;
        /// <summary>
        /// 첫 MouseUp 이벤트 발생 히 자동 전체선택 활성화
        /// </summary>
        /// <param name="edit">설정할 컨트롤</param>
        public void EnableAutoSelectAllOnFirstMouseUp(TextEdit edit)
        {
            edit.MaskBox.MouseUp += MaskBox_MouseUp;
            edit.MaskBox.Enter += MaskBox_Enter;
        }

        /// <summary>
        /// 마스크박스 진입 시 컨트롤 저장
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MaskBox_Enter(object sender, EventArgs e)
        {
            lastEdit = sender;
        }

        /// <summary>
        /// 마스크박스 MouseUp 시 진입했을 경우 전체선택 후 저장 컨트롤 삭제
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void MaskBox_MouseUp(object sender, MouseEventArgs e)
        {
            if (lastEdit == sender)
                (sender as TextBox).SelectAll();
            lastEdit = null;
        }
    }
}