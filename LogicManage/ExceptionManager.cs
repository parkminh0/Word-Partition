using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordPartition
{
    class ExceptionManager : ApplicationException
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////// Field
        ////////////////////////////////////////////////////////////////////////////////////////// Private

        #region Field

        /// <summary>
        /// 네임스페이스
        /// </summary>
        private string _strNamespace;

        /// <summary>
        /// 클래스
        /// </summary>
        private string _strClass;

        /// <summary>
        /// 메소드
        /// </summary>
        private string _strMethod;

        /// <summary>
        /// 에러 코드
        /// </summary>
        private int _nErrorCode;

        /// <summary>
        /// 설명
        /// </summary>
        private string _strDescription;

        /// <summary>
        /// 예외
        /// </summary>
        private Exception _pException;

        /// <summary>
        /// 생성 일시
        /// </summary>
        private DateTime _dteDateTimeCreated;

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Property
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region Property

        #region 네임스페이스 - Namespace

        /// <summary>
        /// 네임스페이스
        /// </summary>
        public string Namespace
        {
            get
            {
                return _strNamespace;
            }
        }

        #endregion

        #region 클래스 - Class

        /// <summary>
        /// 클래스
        /// </summary>
        public string Class
        {
            get
            {
                return _strClass;
            }
        }

        #endregion

        #region 메소드 - Method

        /// <summary>
        /// 메소드
        /// </summary>
        public string Method
        {
            get
            {
                return _strMethod;
            }
        }

        #endregion

        #region 에러 코드 - ErrorCode

        /// <summary>
        /// 에러 코드
        /// </summary>
        public int ErrorCode
        {
            get
            {
                return _nErrorCode;
            }
        }

        #endregion

        #region 설명 - Description

        /// <summary>
        /// 설명
        /// </summary>
        public string Description
        {
            get
            {
                return _strDescription;
            }
        }

        #endregion

        #region 예외 - Exception

        /// <summary>
        /// 예외
        /// </summary>
        public Exception Exception
        {
            get
            {
                return _pException;
            }
        }

        #endregion

        #region 생성 일시 - DateTimeCreated

        /// <summary>
        /// 생성 일시
        /// </summary>
        public DateTime DateTimeCreated
        {
            get
            {
                return _dteDateTimeCreated;
            }
        }

        #endregion

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Constructor
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, nErrorCode, strDescription, pException, dteDateTimeCreated)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        /// <param name="dteDateTimeCreated">생성 일시</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, int nErrorCode, string strDescription, Exception pException, DateTime dteDateTimeCreated)
        {
            _strNamespace = strNamespace;
            _strClass = strClass;
            _strMethod = strMethod;
            _nErrorCode = nErrorCode;
            _strDescription = strDescription;
            _pException = pException;
            _dteDateTimeCreated = dteDateTimeCreated;
        }

        #endregion

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, nErrorCode, strDescription, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스/param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, int nErrorCode, string strDescription, Exception pException)
            : this(strNamespace, strClass, strMethod, nErrorCode, strDescription, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, nErrorCode, strDescription)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스/param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, int nErrorCode, string strDescription)
            : this(strNamespace, strClass, strMethod, nErrorCode, strDescription, null, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, nErrorCode, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스/param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, int nErrorCode, Exception pException)
            : this(strNamespace, strClass, strMethod, nErrorCode, null, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, strDescription, pException, dteDateTimeCreated)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스/param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        /// <param name="dteDateTimeCreated">생성 일시</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, string strDescription, Exception pException, DateTime dteDateTimeCreated)
            : this(strNamespace, strClass, strMethod, -1, strDescription, pException, dteDateTimeCreated)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, strDescription, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스/param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, string strDescription, Exception pException)
            : this(strNamespace, strClass, strMethod, -1, strDescription, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, strDescription)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스/param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, string strDescription)
            : this(strNamespace, strClass, strMethod, -1, strDescription, null, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(strNamespace, strClass, strMethod, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="strNamespace">네임스페이스</param>
        /// <param name="strClass">클래스/param>
        /// <param name="strMethod">메소드</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(string strNamespace, string strClass, string strMethod, Exception pException)
            : this(strNamespace, strClass, strMethod, -1, null, pException, DateTime.Now)
        {
        }

        #endregion


        #region 생성자 - ExceptionManager(pSender, strMethod, nErrorCode, strDescription, pException, dteDateTimeCreated)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        /// <param name="dteDateTimeCreated">생성 일시</param>
        public ExceptionManager(object pSender, string strMethod, int nErrorCode, string strDescription, Exception pException, DateTime dteDateTimeCreated)
        {
            _strNamespace = pSender.GetType().Namespace;
            _strClass = pSender.GetType().Name;
            _strMethod = strMethod;
            _nErrorCode = nErrorCode;
            _strDescription = strDescription;
            _pException = pException;
            _dteDateTimeCreated = dteDateTimeCreated;
        }

        #endregion

        #region 생성자 - ExceptionManager(pSender, strMethod, nErrorCode, strDescription, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(object pSender, string strMethod, int nErrorCode, string strDescription, Exception pException)
            : this(pSender, strMethod, nErrorCode, strDescription, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pSender, strMethod, nErrorCode, strDescription)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        public ExceptionManager(object pSender, string strMethod, int nErrorCode, string strDescription)
            : this(pSender, strMethod, nErrorCode, strDescription, null, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pSender, strMethod, nErrorCode, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(object pSender, string strMethod, int nErrorCode, Exception pException)
            : this(pSender, strMethod, nErrorCode, null, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pSender, strMethod, strDescription, pException, dteDateTimeCreated)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        /// <param name="dteDateTimeCreated">생성 일시</param>
        public ExceptionManager(object pSender, string strMethod, string strDescription, Exception pException, DateTime dteDateTimeCreated)
            : this(pSender, strMethod, -1, strDescription, pException, dteDateTimeCreated)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pSender, strMethod, strDescription, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(object pSender, string strMethod, string strDescription, Exception pException)
            : this(pSender, strMethod, -1, strDescription, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pSender, strMethod, strDescription)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        public ExceptionManager(object pSender, string strMethod, string strDescription)
            : this(pSender, strMethod, -1, strDescription, null, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pSender, strMethod, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pSender">예외 발생자</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(object pSender, string strMethod, Exception pException)
            : this(pSender, strMethod, -1, null, pException, DateTime.Now)
        {
        }

        #endregion


        #region 생성자 - ExceptionManager(pType, strMethod, nErrorCode, strDescription, pException, dteDateTimeCreated)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">타입</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        /// <param name="dteDateTimeCreated">생성 일시</param>
        public ExceptionManager(Type pType, string strMethod, int nErrorCode, string strDescription, Exception pException, DateTime dteDateTimeCreated)
        {
            _strNamespace = pType.Namespace;
            _strClass = pType.Name;
            _strMethod = strMethod;
            _nErrorCode = nErrorCode;
            _strDescription = strDescription;
            _pException = pException;
            _dteDateTimeCreated = dteDateTimeCreated;
        }

        #endregion

        #region 생성자 - ExceptionManager(pType, strMethod, nErrorCode, strDescription, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">타입</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(Type pType, string strMethod, int nErrorCode, string strDescription, Exception pException)
            : this(pType, strMethod, nErrorCode, strDescription, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pType, strMethod, nErrorCode, strDescription)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">타입</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="strDescription">설명</param>
        public ExceptionManager(Type pType, string strMethod, int nErrorCode, string strDescription)
            : this(pType, strMethod, nErrorCode, strDescription, null, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pType, strMethod, nErrorCode, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">타입</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="nErrorCode">에러 코드</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(Type pType, string strMethod, int nErrorCode, Exception pException)
            : this(pType, strMethod, nErrorCode, null, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pType, strMethod, strDescription, pException, dteDateTimeCreated)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">종류</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        /// <param name="dteDateTimeCreated">생성 일시</param>
        public ExceptionManager(Type pType, string strMethod, string strDescription, Exception pException, DateTime dteDateTimeCreated)
            : this(pType, strMethod, -1, strDescription, pException, dteDateTimeCreated)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pType, strMethod, strDescription, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">타입</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(Type pType, string strMethod, string strDescription, Exception pException)
            : this(pType, strMethod, -1, strDescription, pException, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pType, strMethod, strDescription)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">타입</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="strDescription">설명</param>
        public ExceptionManager(Type pType, string strMethod, string strDescription)
            : this(pType, strMethod, -1, strDescription, null, DateTime.Now)
        {
        }

        #endregion

        #region 생성자 - ExceptionManager(pType, strMethod, pException)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="pType">타입</param>
        /// <param name="strMethod">메소드</param>
        /// <param name="pException">예외</param>
        public ExceptionManager(Type pType, string strMethod, Exception pException)
            : this(pType, strMethod, -1, null, pException, DateTime.Now)
        {
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////// Private

        #region 생성자 - ExceptionManager()

        /// <summary>
        /// 생성자
        /// </summary>
        private ExceptionManager()
        {
        }

        #endregion
    }
}
