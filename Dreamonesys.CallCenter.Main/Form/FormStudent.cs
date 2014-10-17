//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;

using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Dreamonesys.CallCenter.Main.Class;

namespace Dreamonesys.CallCenter.Main
{
    /// <summary>
    /// 콜센터 학생검색 클래스
    /// </summary>
    /// 
    public partial class FormStudent : Form
    {

         #region Field

        private Common _common;
        private AppMain _appMain;

        #endregion Field

        #region Property

        #endregion Property

        #region Constructor

        /// <summary>
        /// 기본 생성자
        /// </summary>
        public FormStudent()
        {
            InitializeComponent();

            // 공통 모듈 클래스 인스턴스 생성
            _common = new Common();
            // 프로그램 정보 클래스 인스턴스 생성
            _appMain = new AppMain();
            // 공용 모듈에서 프로그램 정보를 참조할 수 있도록 함
            _common._appMain = _appMain;
            // 프로그램 정보에서 메인 폼을 참조할 수 있도록 함
            _appMain.MainForm = this;
        }

        #endregion Constructor
        
        #region Method

        /// <summary>
        /// 콤보박스 리스트를 조회한다.
        /// </summary>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        public void InitCombo()
        {
            // 캠퍼스 구분 콤보박스 데이터 생성
            //_common.GetComboList(comboBoxCampusType, "캠퍼스구분", true);
            // 캠퍼스 콤보박스 데이터 생성
            //_common.GetComboList(comboBoxCampus, "캠퍼스", true);

            // 콤보박스 멀티
            Common.ComboBoxList[] comboBoxList = 
            {
                new Common.ComboBoxList(comboBoxCampusType, "캠퍼스구분", true),
                new Common.ComboBoxList(comboBoxCampus, "캠퍼스", true)                
            };
            this._common.GetComboList(comboBoxList);
        }

        /// <summary>
        /// 사용자 정의 목록을 조회한다.
        /// </summary>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void SelectDataGridView(DataGridView pDataGridView, string pQueryKind)
        {
            SqlCommand sqlCommand = null;
            SqlResult sqlResult = new SqlResult();

            // 그리드 초기화
            switch (pDataGridView.Name)
            {
                case "dataGridViewStudent":
                    dataGridViewEduStudentClass.Rows.Clear();
                    dataGridViewU2mStudentClass.Rows.Clear();
                    dataGridViewCamMember.Rows.Clear();                    
                    break;
                
                default:
                    break;
            }

            this.Cursor = Cursors.WaitCursor;

            try
            {
                CreateSql(ref sqlCommand, pQueryKind);

                // 쿼리실행 -> 결과값 저장
                this._common.Execute(sqlCommand, ref sqlResult);

                // 성공여부 판단
                if (sqlResult.Success == true)
                {
                    //그리드 초기화
                    pDataGridView.Rows.Clear();

                    // 데이터 테이블 행 루프
                    foreach (DataRow row in sqlResult.DataTable.Rows)
                    {
                        // 그리드 행추가
                        pDataGridView.Rows.Add();

                        //pDataGridView[0, pDataGridView.Rows.Count - 1].Value = pDataGridView.Rows.Count - 1;
                        // 컬럼 루프
                        for (int colCount = 0; colCount <= pDataGridView.Columns.Count - 1; colCount++)
                        {
                             pDataGridView[colCount, pDataGridView.Rows.Count - 1].Value =
                                 row[pDataGridView.Columns[colCount].DataPropertyName].ToString();
                           // pDataGridView[pDataGridView.Columns[colCount].DataPropertyName, pDataGridView.Rows.Count - 1].Value = 
                           //row[pDataGridView.Columns[colCount].DataPropertyName].ToString();
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// pSqlCommand 객체에 쿼리를 정의한다.
        /// </summary>
        /// <param name="pSqlCommand">SqlCommand 객체</param>
        /// <param name="pQueryKind">사용할 쿼리 구분</param>
        /// <param name="pParameter">파라미터</param>
        /// <returns>SqlCommand</returns>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private SqlCommand CreateSql(ref SqlCommand pSqlCommand, string pQueryKind, string[] pParameter = null)
        {
            pSqlCommand = new SqlCommand();
            string businessCD = comboBoxCampusType.SelectedValue.ToString();
            string cpno = comboBoxCampus.SelectedValue.ToString();

            switch (pQueryKind)
            {
                case "select_Student":
                    // 학생 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT  C.cpnm
                             , C.business_cd
                             , C.cpno	
                             , C.cpid
                	         , D.login_char
                             , A.userid
                             , A.usernm
                	         , A.member_id
                             , A.login_id
                             , A.login_pwd
                             , A.grade_cd 
					         , A.phone
					         , A.cell
					         , E.pcell
					         , A.use_yn
					         , A.point
                             , F.db_link
                	    FROM tls_member AS A
                   LEFT JOIN tls_cam_member AS B
                	      ON A.userid = B.userid
                   LEFT JOIN tls_campus AS C
                	      ON B.cpno = C.cpno
                  INNER JOIN tls_campus_group AS D
                	      ON C.cp_group_id = D.cp_group_id
				   LEFT JOIN tls_family AS E
					      ON A.userid = E.userid
                   LEFT JOIN tls_campus_group AS F
                          ON C.cp_group_id = F.cp_group_id
                	   WHERE A.auth_cd = 's' 
                         AND (B.edate = '' OR B.edate IS NULL OR B.sdate <= B.edate)
                          ";

                    if (!string.IsNullOrEmpty(businessCD))
                    {
                        pSqlCommand.CommandText += @"
                         AND C.business_cd = '" + businessCD + "' ";
                    }
                    if (!string.IsNullOrEmpty(cpno))
                    {
                        pSqlCommand.CommandText += @"
                         AND C.cpno = '" + cpno + "' ";
                    }
                    if (!string.IsNullOrEmpty(textBoxUserNm.Text))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.usernm LIKE '%" + textBoxUserNm.Text + "%' ";
                    }                   

                    pSqlCommand.CommandText += @"
                      ORDER BY A.USE_YN DESC, C.CPNM, A.USERNM ";
                    break;

                case "select_edu_student_class":
                    //드림+ 학생 반 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT DISTINCT 
			                  USC.class_id
			                , TC.clno
  			                , TC.clnm
			                , USC.start_date
			                , USC.end_date				
			                , TC.wk_day
			                , (SELECT usernm FROM tls_member WHERE userid = TC.class_tid) AS CLASS_TID
			                , (SELECT name FROM tls_web_code WHERE cdmain = 'grade' AND cdsub = TC.grade_cd) AS GRADE_NM
			                , (SELECT COUNT(userid)FROM tls_class_user WHERE clno = TC.clno AND cpno = TC.cpno AND auth_cd = 's'
				                  AND (end_date = '' OR end_date IS NULL OR CONVERT(CHAR,GETDATE(),112) BETWEEN start_date AND end_date)) AS USER_CNT
		                FROM " + GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "db_link") + @".DBO.V_u2m_student_class AS USC WITH(nolock)
                   LEFT JOIN tls_class as TC 
	    	              ON USC.class_id = TC.class_id
    	               WHERE USC.student_id = '" + GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "member_id") + @"'
		                 AND TC.cpno = " + GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "cpno") + @"
		               ORDER BY USC.end_date ASC
                    ";
                    break;

                case "select_u2m_student_class":
                    //U2M 학생 반 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT TC.class_id
               	            , TCU.clno
                            , TC.clnm
                            , TCU.start_date
                            , TCU.end_date
                            , TC.wk_day
                            , TM.usernm
                            , (SELECT name FROM tls_web_code WHERE cdmain = 'grade' AND cdsub = TC.grade_cd) AS GRADE_NM
               	            , (SELECT COUNT(userid) 
                                 FROM tls_class_user 
                                WHERE clno = TCU.clno AND auth_cd = 's'
								  AND (end_date = '' OR end_date IS NULL OR CONVERT(CHAR,GETDATE(),112) BETWEEN start_date AND end_date)) AS USER_CNT
               	        FROM tls_class_user AS TCU
                   LEFT JOIN tls_class AS TC
              	          ON TCU.clno = TC.clno
                   LEFT JOIN tls_member AS TM
               	          ON TC.class_tid = TM.userid
                       WHERE TCU.userid = " + GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "userid") + @"
                         AND (TCU.end_date = '' OR TCU.end_date IS NULL OR TCU.start_date <= TCU.end_date)
               	       ORDER BY TCU.end_date ASC
                    ";
                    break;          
       
                case "select_cam_member":
                    //캠퍼스 멤버 조회
                    pSqlCommand.CommandText = @"
                      	select REPLACE(b.cpnm, '캠퍼스', '') AS CPNM
			                 , B.cpno
			                 , A.userid
			                 , A.sdate
			                 , A.edate
			                 , A.udatetime
		                  FROM tls_cam_member AS A
	                 LEFT JOIN tls_campus AS B
			                ON A.cpno = B.cpno
	                 LEFT JOIN tls_member AS C
	                        ON A.userid = C.userid
		                 WHERE A.userid = " + GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "userid") + @"
                    ";
                    break;


                default:
                    break;
            }

            return pSqlCommand;
        }

        /// <summary>
        /// 그리드에서 특정 행렬의 값을 리턴한다.
        /// </summary>
        /// <param name="pDataGridView">그리드</param>
        /// <param name="pRowIndex">행번호</param>
        /// <param name="pDataPropertyName">컬럼에 바인딩된 데이터베이스 열</param>
        /// <returns>그리드에서 특정 행렬의 값</returns>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private string GetCellValue(DataGridView pDataGridView, int pRowIndex, string pDataPropertyName = "")
        {
            string CellValue = "";

            foreach (DataGridViewColumn item in pDataGridView.Columns)
            {
                if (item.DataPropertyName.ToLower() == pDataPropertyName)
                {
                    CellValue = pDataGridView[item.Index, pRowIndex].Value.ToString();
                    break;
                }
            }

            return CellValue;
        }

        #endregion Method

        /// <summary>
        /// 폼 로드
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void FormStudent_Load(object sender, EventArgs e)
        {
            InitCombo();
           
        }


        /// <summary>
        /// 캠퍼스 구분 콤보박스 선택 변경시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void comboBoxCampusType_SelectedIndexChanged(object sender, EventArgs e)
        {
             //캠퍼스 콤보박스 데이터 생성
            string campusType = comboBoxCampusType.SelectedValue.ToString();

            _common.GetComboList(comboBoxCampus, "캠퍼스", true, new string[] { campusType });
            //SelectDataGridView(dataGridViewStudent, "select_Student");
        }

        /// <summary>
        /// 캠퍼스 콤보박스 선택 변경시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void comboBoxCampus_SelectedIndexChanged(object sender, EventArgs e)
        {
            //SelectDataGridView(dataGridViewStudent, "select_Student");
        }

        private void textBoxUserNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //학생을 검색한다.                
                SelectDataGridView(dataGridViewStudent, "select_Student");
            }
        }

        
        private void dataGridViewStudent_Click(object sender, EventArgs e)
        {
            if (dataGridViewStudent.Rows.Count > 0 && dataGridViewStudent.CurrentCell != null)
            {
                //드림+ 학생 반 목록을 조회한다.
                SelectDataGridView(dataGridViewEduStudentClass, "select_edu_student_class");
                //U2M 학생 반 목록을 조회한다.
                SelectDataGridView(dataGridViewU2mStudentClass, "select_u2m_student_class");
                //캠퍼스 멤버 조회
                SelectDataGridView(dataGridViewCamMember, "select_cam_member");
                // userid, login_id, login_pw 텍스트박에서 표시한다.  
                textBoxUserid.Text = GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "userid");
                textBoxLoginID.Text = GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "login_id");
                textBoxLoginPW.Text = GetCellValue(dataGridViewStudent, dataGridViewStudent.CurrentCell.RowIndex, "login_pwd");                
<<<<<<< HEAD
            }            
=======
            }
                        
        }

        private void dataGridViewStudent_MouseClick(object sender, MouseEventArgs e)
        {
            //학생 u2m학습창 및 마이페이지 로그인
            if (e.Button == MouseButtons.Right)
            {
                int currentMouseOverRow = ((DataGridView)sender).HitTest(e.X, e.Y).RowIndex;
                if (currentMouseOverRow >= 0)
                {
                    ((DataGridView)sender).CurrentCell = ((DataGridView)sender)[0, currentMouseOverRow];
                    this._common.RunLogin(((DataGridView)sender), new Point(e.X, e.Y));
                }
            }





>>>>>>> 601a29752154d0b7db5ca3f4ab9673dfe17fcee6
        }

        private void dataGridViewStudent_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int currentMouseOverRow = ((DataGridView)sender).HitTest(e.X, e.Y).RowIndex;
                if (currentMouseOverRow >= 0)
                {
                    ((DataGridView)sender).CurrentCell = ((DataGridView)sender)[0, currentMouseOverRow];
                    this._common.RunLogin(((DataGridView)sender), new Point(e.X, e.Y));
                }
            }
        }
       
        #region Event

        

        #endregion Event


    }
}
