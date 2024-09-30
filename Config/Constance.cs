using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace WordPartition
{
    public class Constance
    {
        //public string ServerIP = "localhost";
        public string ServerIP = "192.168.35.188";
        public string compName = "wooribnc";
        public string SystemTitle = "WordPartition";
        //-----------------------------------------------------------------------------------------------------------------//
        // Program Update History
        // 2015.01.05 (v1.0) : 최초 신규 
        // 2018.11.08
        // DBMSSQL Scalar 메소드 내 try/catch 및 Close 메소드 추가
        // DBMSSQL ExcuteTransaction 메소드 Return 형식 ex.Message 문자열로 변경(공백시 오류없음)
        // YmDateEdit 커스텀 컨트롤 추가 (yyyy년 MM월) 형태의 기본 설정 DateEdit
        // 2021.03.18
        // DBManager 클래스 기본 DB를 SQLite에서 MSSQL로 변경
        // MSSQL BulkCopyDI 메소드 추가
        // 기본 폰트 맑은 고딕으로 변경
        // 메인 화면 프로그램 정보, 종료 아이콘 2D로 변경
        //-----------------------------------------------------------------------------------------------------------------//
        public string DBTestQuery = "select CURRENT_DATE";
        public bool DBConnectInSplash = true;
        
        /// <summary>
        /// OS공통 작업폴더
        /// </summary>
        public string CommonFilePath
        {
            get
            {
                return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            }
        }

        /// <summary>
        /// DB FILE위치
        /// </summary>
        private string dbFilePathOrg = Environment.CurrentDirectory;
        public string DbFilePathOrg
        {
            get { return dbFilePathOrg; }
        }

        //private string dbFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "wooribnc");
        private string dbFilePath;
        /// <summary>
        /// 실제 DB파일이 위치할 폴더
        /// </summary>
        public string DbFilePath
        {
            get { return dbFilePath; }
            set
            {
                dbFilePath = value;
            }
        }

        private string dbFileName = "ALCStaff.db"; // SQLite
        public string DbFileName
        {
            get { return dbFileName; }
        }

        /// <summary>
        /// 실제 DB파일의 FULL PATH + Name
        /// </summary>
        public string DbTargetFullName
        {
            get
            {
                return Path.Combine(dbFilePath, dbFileName);
            }
        }

        private string dbBaseFileName = "ALCStaff_Initial.db";
        private string dbInitFileName = "ALCStaff_Initial.db";
        /// <summary>
        /// Initial DB Name
        /// </summary>
        public string DbInitFileName
        {
            get { return dbInitFileName; }
        }
        /// <summary>
        /// 원본 Init파일 Full Path + Name
        /// </summary>
        public string DbInitOrgFullName
        {
            get
            {
                return Path.Combine(dbFilePathOrg, dbInitFileName);
            }
        }

        /// <summary>
        /// MainDB가 되기 직전의 Full Name
        /// </summary>
        public string DbInitFullName
        {
            get
            {
                return Path.Combine(dbFilePath, dbInitFileName);
            }
        }

        /// <summary>
        /// 원본 base파일 Full Path + Name : 초기데이터 들어있는 Init DB
        /// </summary>
        public string DbBaseOrgFullName
        {
            get
            {
                return Path.Combine(dbFilePathOrg, dbBaseFileName);
            }
        }

    }
}
