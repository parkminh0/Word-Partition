using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraEditors;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars.Ribbon;
using DevExpress.Utils.Extensions;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.CodeParser;
using System.IO;
using DevExpress.XtraRichEdit;
using System.Xml;
using System.Text.RegularExpressions;
using DevExpress.Charts.Native;
using DevExpress.XtraRichEdit.Commands;

namespace WordPartition
{
    public partial class MainForm : RibbonForm
    {
        string originalFileName;
        string originalFilePath;
        private DataTable dtBookMark;
        private bool EditMode = false;
        const int colorLength = 15;
        RichEditControl tempREC = new RichEditControl();
        RichEditControl printREC = new RichEditControl();
        private string path = "";

        #region 메인폼
        /// <summary>
        /// 
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 폼 로드
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.UserSkin != "")
                defaultLookAndFeel1.LookAndFeel.SkinName = Properties.Settings.Default.UserSkin;

            dtBookMark = BookMark.SetdtBookMark();
        }

        /// <summary>
        /// 폼 쇼운
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Shown(object sender, EventArgs e)
        {
            Application.DoEvents();
        }

        /// <summary>
        /// 프로그램 종료시
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            if (XtraMessageBox.Show("프로그램을 종료하시겠습니까?", "종료", MessageBoxButtons.YesNo, MessageBoxIcon.Information) != System.Windows.Forms.DialogResult.Yes)
                e.Cancel = true;
            else
                Program.SaveConfig();
        }

        /// <summary>
        /// 프로그램 종료
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bbtnExit_ItemClick(object sender, ItemClickEventArgs e)
        {
            Close();
        }
        #endregion

        #region doc 파일 로드
        /// <summary>
        /// doc 파일 업로드
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bbtnLoad_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.isBusy == true)
                return;

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Word(*.docx;*.doc)|*.docx;.doc"; 

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    originalFilePath = dialog.FileName;
                    if (originalFilePath[originalFilePath.Length - 1] == 'x') //docx
                        originalFileName = dialog.SafeFileName.Substring(0, dialog.SafeFileName.Length  - 5);
                    else //doc
                        originalFileName = dialog.SafeFileName.Substring(0, dialog.SafeFileName.Length  - 4);

                    SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
                    SplashScreenManager.Default.SetWaitFormDescription("파일 로딩중.."); 

                    recOriginal.LoadDocument(originalFilePath);
                    tempREC.Document.RtfText = recOriginal.Document.RtfText;
                    printREC.Document.RtfText = recOriginal.Document.RtfText;
                    ClearBookMark(false); // 파일 로드
                    spnNum.EditValue = 1;
                    rdoDivMode.SelectedIndex = 0;
                    groupControl1.Text = $"≡ 분할지점(총 글자수 {string.Format("{0:n0}", recOriginal.Document.Length)})";
                    groupControl2.Text = "";

                    SplashScreenManager.CloseForm(false);
                }
            }
        }
        #endregion

        #region 분할 및 북마크 생성
        /// <summary>
        /// 분할 시작
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPartition_Click(object sender, EventArgs e)
        {
            // 파일 존재 여부 확인
            DocumentRange tempRange = recOriginal.Document.CreateRange(0, 1);
            if (string.IsNullOrWhiteSpace(recOriginal.Document.GetText(tempRange)))
            {
                XtraMessageBox.Show("파일을 로드한 후 시도해주세요.", "분할", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 유효성 검사
            int setNum = (int)spnNum.Value;
            switch (rdoDivMode.SelectedIndex)
            {
                case 0: // 파일수
                    if (recOriginal.Document.Length / setNum < colorLength + 1)
                    {
                        XtraMessageBox.Show("각 구간의 길이는 16 이상이어야 합니다.", "분할 설정", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    break;
                case 1: // 글자수
                    if (setNum < colorLength + 1)
                    {
                        XtraMessageBox.Show("각 구간의 길이는 16 이상이어야 합니다.", "분할 설정", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        return;
                    }
                    break;
            }

            SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
            SplashScreenManager.Default.SetWaitFormDescription("파일 분할중..");
            
            recOriginal.Document.RtfText = tempREC.Document.RtfText;
            ClearColor();
            ClearBookMark(false); // 파일 분할
            CreateBookMark();
            SetColor();
            SplashScreenManager.CloseForm(false);
        }

        /// <summary>
        /// 북마크, 데이터테이블 초기화
        /// </summary>
        /// <param name="isEdit">true: 수정모드</param>
        private void ClearBookMark(bool isEdit)
        {
            if (!isEdit)
                dtBookMark.Rows.Clear();

            // 북마크 삭제
            for (int i = recOriginal.Document.Bookmarks.Count - 1; i >= 0; i--)
                recOriginal.Document.Bookmarks.Remove(recOriginal.Document.Bookmarks[i]);
           
            for (int i = tempREC.Document.Bookmarks.Count - 1; i >= 0; i--)
                tempREC.Document.Bookmarks.Remove(tempREC.Document.Bookmarks[i]);
            
            for (int i = printREC.Document.Bookmarks.Count - 1; i >= 0; i--)
                printREC.Document.Bookmarks.Remove(printREC.Document.Bookmarks[i]);
        }

        /// <summary>
        /// 북마크 생성
        /// </summary>
        /// <param name="setNum">사용자 입력값</param>
        /// <param name="partitionMode">0: 파일수분할 1: 글자수분할</param>
        private void CreateBookMark()
        {
            int setNum = (int)spnNum.Value;
            // 파일수 분할일 경우 setNum 재설정
            if (rdoDivMode.SelectedIndex == 0)
            {
                double temp2 = (double)recOriginal.Document.Length / (double)setNum;
                setNum = (int)Math.Ceiling(temp2);
            }

            int start = 0;
            int bookmarkName = 1;
            while (start < recOriginal.Document.Length)
            {
                // 북마크 생성
                DocumentRange documentRange;
                if (start + setNum <= recOriginal.Document.Length)
                    documentRange = recOriginal.Document.CreateRange(start, setNum);
                else
                    documentRange = recOriginal.Document.CreateRange(start, recOriginal.Document.Length - start);

                if (documentRange.Length == 0)
                    break;

                recOriginal.Document.Bookmarks.Create(documentRange, bookmarkName.ToString());
                int end = int.Parse(documentRange.End.ToString());

                // dtBookMark 생성
                DocumentRange startString = recOriginal.Document.CreateRange(start, colorLength);
                DocumentRange endString = recOriginal.Document.CreateRange(end - colorLength, colorLength);

                DataRow dr = dtBookMark.NewRow();
                dr[0] = bookmarkName;
                dr[1] = start; // 시작점
                dr[2] = end; // 끝점
                dr[3] = recOriginal.Document.GetText(startString).Replace("\r", "").Replace("\n", ""); // 시작 텍스트
                dr[4] = recOriginal.Document.GetText(endString).Replace("\r", "").Replace("\n", ""); // 끝 텍스트
                dr[5] = end - start;
                dtBookMark.Rows.Add(dr);

                start += setNum;
                bookmarkName++;
            }

            // gridControl 
            grdBookMark.DataSource = dtBookMark;
            grdViewBookMark.BestFitColumns();
        }
        #endregion

        #region 북마크 색칠
        /// <summary>
        /// 색칠 지우기
        /// </summary>
        private void ClearColor()
        {
            foreach (Bookmark bm in recOriginal.Document.Bookmarks)
            {
                // 시작 string
                DocumentRange startRange = recOriginal.Document.CreateRange(bm.Range.Start.ToInt(), colorLength);
                CharacterProperties setStart = recOriginal.Document.BeginUpdateCharacters(startRange);
                setStart.HighlightColor = Color.White;
                setStart.ForeColor = Color.Black;
                recOriginal.Document.EndUpdateCharacters(setStart);

                // 끝 string
                DocumentRange endRange = recOriginal.Document.CreateRange(bm.Range.End.ToInt() - colorLength, colorLength);
                CharacterProperties setEnd = recOriginal.Document.BeginUpdateCharacters(endRange);
                setEnd.HighlightColor = Color.White;
                setEnd.ForeColor = Color.Black;
                recOriginal.Document.EndUpdateCharacters(setEnd);
            }
        }

        /// <summary>
        /// 색칠하기
        /// </summary>
        private void SetColor()
        {
            foreach (Bookmark bm in recOriginal.Document.Bookmarks)
            {
                // 시작 string
                DocumentRange startRange = recOriginal.Document.CreateRange(bm.Range.Start.ToInt(), colorLength);
                CharacterProperties setStart = recOriginal.Document.BeginUpdateCharacters(startRange);
                setStart.Reset();
                setStart.HighlightColor = Color.FromArgb(249, 235, 222);
                setStart.ForeColor = Color.FromArgb(129, 88, 84);
                recOriginal.Document.EndUpdateCharacters(setStart);

                // 끝 string
                DocumentRange endRange = recOriginal.Document.CreateRange(bm.Range.End.ToInt() - colorLength, colorLength);
                CharacterProperties setEnd = recOriginal.Document.BeginUpdateCharacters(endRange);
                setEnd.Reset();
                setEnd.HighlightColor = Color.FromArgb(129, 88, 84);
                setEnd.ForeColor = Color.FromArgb(249, 235, 222);
                recOriginal.Document.EndUpdateCharacters(setEnd);
            }
        }
        #endregion

        #region 북마크 이동, 북마크 수정
        /// <summary>
        /// 북마크 이동
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void grdViewBookMark_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            DataRow row = grdViewBookMark.GetFocusedDataRow();
            if (row == null)
            {
                groupControl2.Text = "";
                return;
            }

            if (EditMode)
                groupControl2.Text = "[수정모드]   " + $"[{row[0]}번 구간]          시작문장 = {row[3]}          끝문장 = {row[4]}";
            else
                groupControl2.Text = $"[{row[0]}번 구간]          시작문장 = {row[3]}          끝문장 = {row[4]}";
            DocumentPosition documentPosition = recOriginal.Document.CreatePosition(int.Parse(row[2].ToString()));
            recOriginal.Document.CaretPosition = documentPosition;
            recOriginal.ScrollToCaret();
        }

        /// <summary>
        /// 포지션 가져오기 및 데이터 그리드 수정
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void recOriginal_Click(object sender, EventArgs e)
        {
            DataRow row = grdViewBookMark.GetFocusedDataRow();
            if (row == null || !EditMode)
                return;

            DocumentPosition documentPosition = recOriginal.Document.CaretPosition; // 현재 선택 위치
            DocumentRange endString = recOriginal.Document.CreateRange(documentPosition.ToInt() - colorLength, colorLength);

            row[2] = documentPosition.ToInt(); // 위치
            row[4] = recOriginal.Document.GetText(endString).Replace("\r", "").Replace("\n", ""); // 북마크 끝부분 텍스트
            row[5] = documentPosition.ToInt() - int.Parse(row[1].ToString());

            if (int.Parse(row[0].ToString()) != dtBookMark.Rows.Count) // 선택한 북마크 바로 다음의 북마크에 영향을 주기 때문에 마지막 북마크가 아닌지 확인
            {
                DocumentRange startString = recOriginal.Document.CreateRange(documentPosition.ToInt(), colorLength);
                dtBookMark.Rows[int.Parse(row[0].ToString())][1] = row[2]; // 시작부분 위치
                dtBookMark.Rows[int.Parse(row[0].ToString())][3] = recOriginal.Document.GetText(startString).Replace("\r", "").Replace("\n", ""); // 시작부분 텍스트
                dtBookMark.Rows[int.Parse(row[0].ToString())][5] = int.Parse(dtBookMark.Rows[int.Parse(row[0].ToString())][2].ToString()) - documentPosition.ToInt();
            }

            grdBookMark.DataSource = dtBookMark;
        }

        /// <summary>
        /// 수정/적용/취소
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void groupControl2_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            if (e.Button.Properties.Tag.ToString() == "Edit") // 수정
                edit();
        }

        /// <summary>
        /// 수정(F5) / 적용(F6) / 취소(F7)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F5: // 수정
                    edit();
                    break;
                case Keys.F6: // 적용
                    apply();
                    break;
                case Keys.F7: // 취소
                    cancel();
                    break;
            }
        }

        /// <summary>
        /// 수정
        /// </summary>
        private void edit()
        {
            if (dtBookMark.Rows.Count == 0 || dtBookMark == null)
                return;

            if (EditMode)
            {
                for (int i = 0; i < dtBookMark.Rows.Count; i++)
                {
                    if (int.Parse(dtBookMark.Rows[i][2].ToString()) - int.Parse(dtBookMark.Rows[i][1].ToString()) < colorLength)
                    {
                        XtraMessageBox.Show("각 구간의 길이는 15 이상이어야 합니다.", "적용", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
                SplashScreenManager.Default.SetWaitFormDescription("수정 사항 적용중..");

                ClearColor();
                ClearBookMark(true); // 북마크 수정
                EditBookMark();
                SetColor();
                DataRow row = grdViewBookMark.GetFocusedDataRow();
                groupControl2.Text = $"[{row[0]}번 구간]    시작문장 = {row[3]}          끝문장 = {row[4]}";
                EditMode = false;

                SplashScreenManager.CloseForm(false);
            }
            else
            {
                groupControl2.Text = "[수정모드]   " + groupControl2.Text;
                EditMode = true;
            }
        }

        /// <summary>
        /// 북마크 업데이트
        /// </summary>
        private void EditBookMark()
        {
            foreach (DataRow dr in dtBookMark.Rows)
            {
                DocumentRange documentRange = recOriginal.Document.CreateRange(int.Parse(dr[1].ToString()), int.Parse(dr[2].ToString()) - int.Parse(dr[1].ToString()));
                recOriginal.Document.Bookmarks.Create(documentRange, dr[0].ToString());
            }
        }
        #endregion

        #region 저장
        /// <summary>
        /// DOCX 저장
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bbtnSave_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (dtBookMark.Rows.Count == 0 || recOriginal.Document.Length == 0)
                return;

            XtraFolderBrowserDialog dialog = new XtraFolderBrowserDialog();
            dialog.Title = "폴더 경로 지정";
            if (dialog.ShowDialog() == DialogResult.OK)
                path = dialog.SelectedPath;
            else
                return;

            SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);

            double num = 0;
            int end = recOriginal.Document.Bookmarks.Count;
            foreach (Bookmark bm in recOriginal.Document.Bookmarks)
            {
                num++;
                SplashScreenManager.Default.SetWaitFormDescription($"파일 저장 중.. {num}/{end}");

                DocumentRange print = printREC.Document.CreateRange(0, printREC.Document.Length);
                printREC.Document.Delete(print);

                DocumentRange documentRange = tempREC.Document.CreateRange(bm.Range.Start, bm.Range.End.ToInt() - bm.Range.Start.ToInt());
                printREC.Document.AppendRtfText(tempREC.Document.GetRtfText(documentRange));
                string newName = $"{originalFileName}_{num}.docx";
                string filepath = GetAvailablePathname(path, newName);
                printREC.SaveDocument(filepath, DocumentFormat.OpenXml);
            }

            SplashScreenManager.CloseForm(false);
            XtraMessageBox.Show("저장이 완료되었습니다.", "저장", MessageBoxButtons.OK);
        }

        /// <summary>
        /// 중복 파일명 확인
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        private string GetAvailablePathname(string folderPath, string filename)
        {
            int invalidChar = 0;
            do
            {
                invalidChar = filename.IndexOfAny(Path.GetInvalidFileNameChars());

                if (invalidChar != -1)
                    filename = filename.Remove(invalidChar, 1);
            }
            while (invalidChar != -1);

            string fullPath = Path.Combine(folderPath, filename);
            string filenameWithoutExtention = Path.GetFileNameWithoutExtension(filename);
            string extension = Path.GetExtension(filename);

            while (File.Exists(fullPath))
            {
                Regex rg = new Regex(@".*\((?<Num>\d*)\)");
                System.Text.RegularExpressions.Match mt = rg.Match(fullPath);

                if (mt.Success)
                {
                    string numberOfCopy = mt.Groups["Num"].Value;
                    int nextNumberOfCopy = int.Parse(numberOfCopy) + 1;
                    int posStart = fullPath.LastIndexOf("(" + numberOfCopy + ")");

                    fullPath = string.Format("{0}({1}){2}", fullPath.Substring(0, posStart), nextNumberOfCopy, extension);
                }
                else
                {
                    fullPath = folderPath + "\\" + filenameWithoutExtention + " (2)" + extension;
                }
            }
            return fullPath;
        }
        #endregion

        #region 추후 사용 가능
        /// <summary>
        /// 북마크 임시 삭제
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Delete_Click(object sender, EventArgs e)
        {
            //if (!EditMode)
            //    return;

            //grdViewBookMark.DeleteSelectedRows();
        }

        /// <summary>
        /// 수정 취소
        /// </summary>
        private void cancel()
        {
            //if (!EditMode)
            //    return;

            //if (XtraMessageBox.Show("초기 상태로 돌아갑니다.\r\n계속하시겠습니까?", "취소", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //{
            //    groupControl2.Text = groupControl2.Text.Substring(9);
            //    grdBookMark.DataSource = dtBookMarkCopy;
            //    dtBookMark = dtBookMarkCopy;
            //    EditMode = false;
            //}
        }

        /// <summary>
        /// 수정내용 적용
        /// </summary>
        private void apply()
        {
            //if (!EditMode || dtBookMark.Rows.Count == 0 || dtBookMark == null)
            //    return;

            //for (int i = 0; i < dtBookMark.Rows.Count; i++)
            //{
            //    if (int.Parse(dtBookMark.Rows[i][2].ToString()) - int.Parse(dtBookMark.Rows[i][1].ToString()) < colorLength)
            //    {
            //        XtraMessageBox.Show("각 구간의 길이는 15 이상이어야 합니다.", "적용", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        return;
            //    }
            //}

            //if (XtraMessageBox.Show("수정 사항을 적용하시겠습니까?", "적용", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //{
            //    SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
            //    SplashScreenManager.Default.SetWaitFormDescription("수정 사항 적용중..");

            //    EditBookMark();
            //    EditMode = false;
            //    DataRow row = grdViewBookMark.GetFocusedDataRow();
            //    groupControl2.Text = $"[{row[0]}번 구간]          시작문장 = {row[3]}          끝문장 = {row[4]}";

            //    SplashScreenManager.CloseForm(false);
            //}
        }
        #endregion
    }
}