using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.UserSkins;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using SecureApp;

namespace WordPartition
{
    static class Program
    {
        public static OptionInfo Option;
        public static WarningMSG WMSG;
        public static bool GRunYN;
        public static bool isBusy;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            constance = new Constance();
            LoadConfig();

            // 기본스킨 Bezier로 변경 및 팔레트 설정 추가 (DB설정 및 로그인 창에도 적용이 되도록 설정시점 변경)
            if (Properties.Settings.Default.UserSkin != "")
            {
                if (Properties.Settings.Default.UserSkin == "The Bezier" && Properties.Settings.Default.UserPalette != "")
                    UserLookAndFeel.Default.SetSkinStyle(Properties.Settings.Default.UserSkin, Properties.Settings.Default.UserPalette);
                else
                    UserLookAndFeel.Default.SetSkinStyle(Properties.Settings.Default.UserSkin);
            }
            else
            {
                UserLookAndFeel.Default.SetSkinStyle("The Bezier", "Gloom Gloom");
            }
            // DB설정 및 로그인 창에도 적용이 되도록 폰트 설정시점 변경
            DevExpress.Utils.AppearanceObject.DefaultFont = new System.Drawing.Font("맑은 고딕", 9);

            GRunYN = true;

            mainApp = new MainForm();
            WMSG = new WarningMSG();
            string runCheck = "";

            if (string.IsNullOrEmpty(runCheck) && GRunYN)
            {
                Application.Run(mainApp);
            }
            else
            {
                Option.DBFullPath = string.Empty;
                SaveConfig();
                WMSG.MSG(runCheck);
            }
        }

        public static MainForm mainApp;
        public static Constance constance;

        /// <summary> 설정파일 xml싱크. </summary>
        public static XmlSerializer m_serializer = new XmlSerializer(typeof(OptionInfo));

        /// <summary>
        /// 설정파일 xml 전체경로
        /// </summary>
        static string cfgPath = Path.GetFileNameWithoutExtension(System.Environment.GetCommandLineArgs()[0]) + "Settings.xml";
        public static string CfgPath
        {
            get
            {
                return Path.Combine(Program.constance.CommonFilePath, cfgPath);
            }
        }

        /// <summary>
        /// 설정정보 로드
        /// </summary>
        public static void LoadConfig()
        {
            if (File.Exists(Program.CfgPath))
            {
                bool isError = false;
                using (FileStream fs = File.OpenRead(Program.CfgPath))
                {
                    try
                    {
                        Program.Option = Program.m_serializer.Deserialize(fs) as OptionInfo;
                    }
                    catch
                    {
                        isError = true;
                    }
                }

                if (isError)
                {
                    File.Delete(Program.CfgPath);
                }
            }

            if (Program.Option == null)
            {
                Program.Option = new OptionInfo();
                Option.LoginID = "admin";
                SaveConfig();
            }
            //DB파일경로 검사
            if (!string.IsNullOrEmpty(Option.DBFullPath))
            {
                try
                {
                    if (Directory.Exists(Option.DBFullPath))
                        constance.DbFilePath = Option.DBFullPath;
                    else
                        Option.DBFullPath = string.Empty;
                }
                catch (Exception)
                {
                    Option.DBFullPath = string.Empty;
                }
                finally
                {
                    SaveConfig();
                }
            }
        }

        /// <summary>
        /// 설정정보 저장.
        /// </summary>
        public static void SaveConfig()
        {
            if (!File.Exists(Program.CfgPath))
            {
                File.Create(Program.CfgPath).Close();
            }

            using (XmlTextWriter xtw = new XmlTextWriter(Program.CfgPath, Encoding.UTF8))
            {
                Program.m_serializer.Serialize(xtw, Program.Option);
                xtw.Flush();
                xtw.Close();
            }
        }
    }
}
