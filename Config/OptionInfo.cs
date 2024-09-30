using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordPartition
{
    public class OptionInfo
    {
        string loginID;
        //string dbConnStr;

        public string LoginID
        {
            get { return loginID; }
            set
            {
                loginID = value;
            }
        }

        private string _DBFullPath;
        /// <summary>
        /// Database File 위치
        /// </summary>
        public string DBFullPath
        {
            get { return _DBFullPath; }
            set { _DBFullPath = value; }
        }

    }
}
