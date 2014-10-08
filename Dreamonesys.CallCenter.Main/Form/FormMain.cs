using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    /// 콜센터 메인 클래스
    /// </summary>
    public partial class FormMain : Form
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
        public FormMain()
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
                //main tab 콤보박스
                new Common.ComboBoxList(comboBoxCampusType, "캠퍼스구분", true),
                new Common.ComboBoxList(comboBoxCampus, "캠퍼스", true),              
                //콩알관리 콤보박스
                new Common.ComboBoxList(comboBoxCampusTypePoint, "캠퍼스구분", true),
                new Common.ComboBoxList(comboBoxCampusPoint, "캠퍼스", true) ,  
                new Common.ComboBoxList(comboBoxSchoolCDPoint, "학교급", true)   
                
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
                case "dataGridViewCampus":
                    dataGridViewClassEmployee.Rows.Clear();
                    dataGridViewEmployee.Rows.Clear();                    
                    break;
                case "dataGridViewEmployee":
                    dataGridViewClassEmployee.Rows.Clear();
                    dataGridViewClassStudent.Rows.Clear();
                    break;
                case "dataGridViewClassEmployee":
                    dataGridViewClassStudent.Rows.Clear();
                    break;
                case "dataGridViewCampusPoint":
                    dataGridViewClassPoint.Rows.Clear();
                    dataGridViewStudentPoint.Rows.Clear();
                    break;
                case "dataGridViewClassPoint":
                    dataGridViewStudentPoint.Rows.Clear();
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
                            if (pDataGridView.Columns[colCount].DataPropertyName != "check_yn")
                            {
                                pDataGridView[colCount, pDataGridView.Rows.Count - 1].Value =
                                    //dataGridViewCampus.Rows[dataGridViewCampus.Rows.Count - 1].Cells[colCount].Value = 
                                    row[pDataGridView.Columns[colCount].DataPropertyName].ToString();                                
                            }
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
            string businessCDPoint = comboBoxCampusTypePoint.SelectedValue.ToString();
            string cpnoPoint = comboBoxCampusPoint.SelectedValue.ToString();
            string schoolCDPoint = comboBoxSchoolCDPoint.SelectedValue.ToString();

            switch (pQueryKind)
            {
                case "select_campus":
                    // 캠퍼스 목록 조회
                    pSqlCommand.CommandText = @"
                     SELECT B.cp_group_nm 
                          , A.cpnm
                		  , CASE A.business_cd WHEN 'DD' THEN '직영'
                                               WHEN 'FA' THEN 'FC'
                                               ELSE 'CP'
                            END business_cd
                		  , A.cpno
                		  , A.cpid
                		  , B.login_char
                		  , B.db_link
                		FROM tls_campus AS A
                   LEFT JOIN tls_campus_group AS B
                		  ON A.cp_group_id = B.cp_group_id
                	   WHERE A.use_yn = 'Y' ";
                    if (!string.IsNullOrEmpty(businessCD))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.business_cd = '" + businessCD + "' ";
                    }
                    if (!string.IsNullOrEmpty(cpno))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.cpno = '" + cpno + "' ";
                    }
                    pSqlCommand.CommandText += @"
                       ORDER BY a.business_cd DESC , B.cp_group_nm DESC, A.cpnm ";
                    break;

                case "select_employee":
                    // 직원 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT A.member_id
                            , A.usernm
                            , A.login_id
                            , A.login_pwd
                            , UE.enter_date
                            , UE.retire_date
                            , A.use_yn
                            , A.tutor_yn
                            , (SELECT name from tls_web_code WHERE cdmain = 'auth' and cdsub = A.auth_cd) AS AUTH_CD
                            , A.userid
                         FROM tls_member AS A
                    LEFT JOIN " + GetCellValue(dataGridViewCampus, dataGridViewCampus.CurrentCell.RowIndex, "db_link") + @".DBO.V_u2m_employee AS UE
                           ON A.member_id = UE.emp_id
                    LEFT JOIN tls_cam_member as B
                           ON A.userid = B.userid
                        WHERE B.cpno = " + GetCellValue(dataGridViewCampus, dataGridViewCampus.CurrentCell.RowIndex, "cpno") + @"
                          AND A.auth_cd <> 'S'
                          AND (UE.retire_date = '' OR UE.retire_date IS NULL) ";
                    if (!string.IsNullOrEmpty(textBoxUserNm.Text))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.usernm LIKE '%" + textBoxUserNm.Text + "%' ";
                    }
                    pSqlCommand.CommandText += @"
                        ORDER BY A.use_yn desc, A.tutor_yn desc, auth_cd, A.usernm
                    ";
                    textBoxUserNm.Text = "";
                    break;

                case "select_class_employee":
                    // 수업교사 반 목록 조회
                    pSqlCommand.CommandText = @"
	                   SELECT B.class_id
	                        , B.clno
	                        , B.clnm
	                        , B.point
	                        , B.mpoint
	  		                , B.school_cd
	                        , C.usernm
                            , A.cpno
	                     FROM tls_class_user AS A 
                    LEFT JOIN tls_class AS B 
	                       ON A.clno = B.clno
                    LEFT JOIN tls_member AS C
	                       ON B.class_tid = C.userid
	                    WHERE A.userid = " + GetCellValue(dataGridViewEmployee, dataGridViewEmployee.CurrentCell.RowIndex, "userid") + @"
                         AND (B.edate = '' OR B.edate IS NULL OR B.edate >= CONVERT(VARCHAR(8), GETDATE(), 112))
                         AND B.use_yn = 'Y'
	                   ORDER BY B.school_cd, B.clnm
                    ";
                    break;

                case "select_class_student":
                    //반 학생 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT A.userid
	                        , C.usernm
                            , A.cpno                            
	                     FROM tls_class_user AS A 
                    LEFT JOIN tls_class AS B 
                           ON A.clno = B.clno
                    LEFT JOIN tls_member AS C
                           ON A.userid = C.userid
                        WHERE B.clno = " + GetCellValue(dataGridViewClassEmployee, dataGridViewClassEmployee.CurrentCell.RowIndex, "clno") + @"
                          AND A.auth_cd = 'S'
                          AND (A.end_date = '' OR A.end_date IS NULL OR A.end_date >= CONVERT(VARCHAR(8), GETDATE(), 112) )
                        ORDER BY C.usernm
                    ";
                    break;

                case "select_campus_point":
                    // 콩알관리 탭 캠퍼스(콩알) 목록 조회
                    pSqlCommand.CommandText = @"
                     SELECT A.cpnm                           
                		  , CASE A.business_cd WHEN 'DD' THEN '직영'
                                               WHEN 'FA' THEN 'FC'
                                               ELSE 'CP'
                            END business_cd
                		  , SUM(C.mpoint) AS POINT 
                          , A.cpno                        
                		FROM tls_campus AS A
                   LEFT JOIN tls_campus_group AS B
                		  ON A.cp_group_id = B.cp_group_id
                   LEFT JOIN tls_class AS C
                          ON A.cpno = C.cpno
                	   WHERE A.use_yn = 'Y' 
                         AND C.use_yn = 'Y'
                         AND (C.edate = '' or C.edate = '' or C.edate >= CONVERT(VARCHAR(8), GETDATE(), 112))
                             ";
                    if (!string.IsNullOrEmpty(businessCDPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.business_cd = '" + businessCDPoint + "' ";
                    }
                    if (!string.IsNullOrEmpty(cpnoPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.cpno = '" + cpnoPoint + "' ";
                    }                    
                    pSqlCommand.CommandText += @"
                       GROUP BY A.cpnm, A.cpno, A.business_cd
                       ORDER BY A.business_cd DESC,  A.cpnm ";
                    break;

                case "select_class_point":
                    //콩알관리 탭 반 콩알 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT clnm
                            , mpoint AS POINT
                            , school_cd
                            , clno
	                     FROM tls_class
                        WHERE cpno = " + GetCellValue(dataGridViewCampusPoint, dataGridViewCampusPoint.CurrentCell.RowIndex, "cpno") + @"
                          AND use_yn = 'Y'
                          AND (edate = '' or edate = '' or edate >= CONVERT(VARCHAR(8), GETDATE(), 112)) ";                    
                    if (!string.IsNullOrEmpty(schoolCDPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND school_cd = '" + schoolCDPoint + "' ";
                    }                    
                    pSqlCommand.CommandText += @"
                       ORDER BY clnm desc ";
                    break;

                case "select_student_point":
                    //해당 반의 학생 콩알 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT D.cpnm
                            , B.clnm
	                        , C.usernm
                            , C.point
                            , A.cpno                            
                            , A.userid
	                     FROM tls_class_user AS A 
                    LEFT JOIN tls_class AS B 
                           ON A.clno = B.clno
                    LEFT JOIN tls_member AS C
                           ON A.userid = C.userid
                    LEFT JOIN tls_campus AS D
                           ON A.cpno = D.cpno
                        WHERE A.cpno = '" + GetCellValue(dataGridViewCampusPoint, dataGridViewCampusPoint.CurrentCell.RowIndex, "cpno") + @"'
                          AND A.clno = '" + GetCellValue(dataGridViewClassPoint, dataGridViewClassPoint.CurrentCell.RowIndex, "clno") + @"'
                          AND A.auth_cd = 'S'
                          AND (A.end_date = '' OR A.end_date IS NULL OR A.end_date >= CONVERT(VARCHAR(8), GETDATE(), 112) ) ";

                     if (!string.IsNullOrEmpty(businessCDPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND D.business_cd = '" + businessCDPoint + "' ";
                    }
                    if (!string.IsNullOrEmpty(cpnoPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.cpno = '" + cpnoPoint + "' ";
                    }
                    if (!string.IsNullOrEmpty(schoolCDPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND B.school_cd = '" + schoolCDPoint + "' ";
                    }       
                    pSqlCommand.CommandText += @"                        
                        ORDER BY D.cpnm, B.clnm, C.usernm ";
                   
                    
                    break;
                case "select_student_point_all":
                    //학생 콩알 목록 조회
                    pSqlCommand.CommandText = @"
                       SELECT D.cpnm
                            , B.clnm
	                        , C.usernm
                            , C.point
                            , A.cpno                            
                            , A.userid
	                     FROM tls_class_user AS A 
                    LEFT JOIN tls_class AS B 
                           ON A.clno = B.clno
                    LEFT JOIN tls_member AS C
                           ON A.userid = C.userid
                    LEFT JOIN tls_campus AS D
                           ON A.cpno = D.cpno
                        WHERE A.auth_cd = 'S'
                          AND (A.end_date = '' OR A.end_date IS NULL OR A.end_date >= CONVERT(VARCHAR(8), GETDATE(), 112) ) ";

                    if (!string.IsNullOrEmpty(businessCDPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND D.business_cd = '" + businessCDPoint + "' ";
                    }
                    if (!string.IsNullOrEmpty(cpnoPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND A.cpno = '" + cpnoPoint + "' ";
                    }
                    if (!string.IsNullOrEmpty(schoolCDPoint))
                    {
                        pSqlCommand.CommandText += @"
                         AND B.school_cd = '" + schoolCDPoint + "' ";
                    }                    
                    
                    if (!string.IsNullOrEmpty(textBoxStudentNMPoint.Text))
                    {
                        pSqlCommand.CommandText += @"
                        AND C.usernm like '%" + textBoxStudentNMPoint.Text + "%' ";
                    }


                    pSqlCommand.CommandText += @"                        
                        ORDER BY D.cpnm, B.clnm, C.usernm ";
                    textBoxStudentNMPoint.Text = "";

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

        /// <summary>
        /// 수업교사 반 목록을 삭제한다.
        /// </summary>
        /// <history>
        /// 박석제, 2014-10-07, 생성
        /// </history>
        private void DeleteClassEmployee()
        {
            SqlCommand sqlCommand = null;
            SqlResult sqlResult = new SqlResult();
            
            this.Cursor = Cursors.WaitCursor;

            
            try
            {

            }
            catch (Exception ex)
            {
                
                throw;
            }
        }

        #endregion Method

        #region Event

        /// <summary>
        /// 폼 로드
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void FormMain_Load(object sender, EventArgs e)
        {
            InitCombo();
            SelectDataGridView(dataGridViewCampus, "select_campus");
        }

        /// <summary>
        /// 캠퍼스 구분 콤보박스 선택 변경시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void comboBoxCampusType_SelectionChangeCommitted(object sender, EventArgs e)
        {
            // 캠퍼스 콤보박스 데이터 생성
            string campusType = comboBoxCampusType.SelectedValue.ToString();

            _common.GetComboList(comboBoxCampus, "캠퍼스", true, new string[] { campusType });
            SelectDataGridView(dataGridViewCampus, "select_campus");
        }

        /// <summary>
        /// 캠퍼스 콤보박스 선택 변경시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void comboBoxCampus_SelectionChangeCommitted(object sender, EventArgs e)
        {
            SelectDataGridView(dataGridViewCampus, "select_campus");
        }

        /// <summary>
        /// 캠퍼스 목록 클릭시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void dataGridViewCampus_Click(object sender, EventArgs e)
        {
            // 교사정보를 조회한다.
            SelectDataGridView(dataGridViewEmployee, "select_employee");
        }

        /// <summary>
        /// 교사정보 클릭시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewEmployee_Click(object sender, EventArgs e)
        {
            if (dataGridViewEmployee.Rows.Count > 0 && dataGridViewEmployee.CurrentCell != null)
            {
                // 수업교사 반 목록을 조회한다.
                SelectDataGridView(dataGridViewClassEmployee, "select_class_employee");
                // userid, login_id, login_pw 텍스트박스에서 표시한다.            
                textBoxUserID.Text = GetCellValue(dataGridViewEmployee, dataGridViewEmployee.CurrentCell.RowIndex, "userid");
                textBoxMemberID.Text = GetCellValue(dataGridViewEmployee, dataGridViewEmployee.CurrentCell.RowIndex, "member_id");
                textBoxLoginPW.Text = GetCellValue(dataGridViewEmployee, dataGridViewEmployee.CurrentCell.RowIndex, "login_pwd");
                
            }
        }

        /// <summary>
        /// 직원명 검색 TextBox 에 Enter 키 입력시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void textBoxUserNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //직원을 검색한다.                
                SelectDataGridView(dataGridViewEmployee, "select_employee");
            }
        }

        private void dataGridViewClassEmployee_Click(object sender, EventArgs e)
        {
            //반 학생 목록을 조회한다.
            SelectDataGridView(dataGridViewClassStudent, "select_class_student");
        }       

        private void toolStripButtonSelect_Student_Click(object sender, EventArgs e)
        {
            //폼2 이동 (학생검색)
            FormStudent frm2 = new FormStudent();
            frm2.Show();
        }

        private void dataGridViewClassEmployee_DoubleClick(object sender, EventArgs e)
        {
            //반 차시 조회 폼 이동
            FormClassSchedule frmSchedule1 = new FormClassSchedule();
            frmSchedule1.ClassEmployeeCPNO = GetCellValue(dataGridViewClassEmployee, dataGridViewClassEmployee.CurrentCell.RowIndex, "cpno");
            frmSchedule1.ClassEmployeeCLNO = GetCellValue(dataGridViewClassEmployee, dataGridViewClassEmployee.CurrentCell.RowIndex, "clno");

            frmSchedule1.Show();            
        }

        private void dataGridViewClassStudent_DoubleClick(object sender, EventArgs e)
        {
            //학생 차시 조회 폼 이동
            FormStudentSchedule frmSchedule2 = new FormStudentSchedule();
            frmSchedule2.ClassStudentCPNO = GetCellValue(dataGridViewClassStudent, dataGridViewClassStudent.CurrentCell.RowIndex, "cpno");            
            frmSchedule2.ClassStudentUID = GetCellValue(dataGridViewClassStudent, dataGridViewClassStudent.CurrentCell.RowIndex, "userid");            
            
            frmSchedule2.Show();
        }

        private void toolStripButtonClassStudy_Click(object sender, EventArgs e)
        {
            //반 차시 조회 폼 이동
            FormClassSchedule frmSchedule1 = new FormClassSchedule();
            frmSchedule1.ClassEmployeeCPNO = GetCellValue(dataGridViewCampus, dataGridViewCampus.CurrentCell.RowIndex, "cpno");            

            frmSchedule1.Show();
        }

        private void toolStripButtonStudentStudy_Click(object sender, EventArgs e)
        {
            //학생 차시 조회 폼 이동
            FormStudentSchedule frmSchedule2 = new FormStudentSchedule();
            frmSchedule2.ClassStudentCPNO = GetCellValue(dataGridViewCampus, dataGridViewCampus.CurrentCell.RowIndex, "cpno");            

            frmSchedule2.Show();
        }

        #endregion Event

        //콩알조회 탭


        /// <summary>
        /// 캠퍼스 구분(포인트) 콤보박스 선택 변경시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void comboBoxCampusTypePoint_SelectionChangeCommitted(object sender, EventArgs e)
        {
            // 캠퍼스 콤보박스 데이터 생성
            string campusType = comboBoxCampusTypePoint.SelectedValue.ToString();

            _common.GetComboList(comboBoxCampusPoint, "캠퍼스", true, new string[] { campusType });
            SelectDataGridView(dataGridViewCampusPoint, "select_campus_point");
        }

        /// <summary>
        /// 캠퍼스(포인트) 콤보박스 선택 변경시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void comboBoxCampusPoint_SelectionChangeCommitted(object sender, EventArgs e)
        {
            SelectDataGridView(dataGridViewCampusPoint, "select_campus_point");
        }

        /// <summary>
        /// 학교급(포인트) 목록 클릭시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void comboBoxSchoolCDPoint_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (dataGridViewCampusPoint.Rows.Count > 0 && dataGridViewCampusPoint.CurrentCell != null)
            {
                SelectDataGridView(dataGridViewClassPoint, "select_class_point");
            }
        }

        /// <summary>
        /// 캠퍼스(포인트) 목록 클릭시 발생하는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성하였음.
        /// </history>
        private void dataGridViewCampusPoint_Click(object sender, EventArgs e)
        {
            //반 콩알정보 조회한다
            if (dataGridViewCampusPoint.Rows.Count > 0 && dataGridViewCampusPoint.CurrentCell != null)
            {
                SelectDataGridView(dataGridViewClassPoint, "select_class_point");
            }
        }

        private void dataGridViewClassPoint_Click(object sender, EventArgs e)
        {
            //학생 콩알정보 조회
            if (dataGridViewClassPoint.Rows.Count > 0 && dataGridViewClassPoint.CurrentCell != null)
            {
                SelectDataGridView(dataGridViewStudentPoint, "select_student_point");
            }
        }

        private void textBoxStudentNMPoint_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //학생콩알정보 조회
                SelectDataGridView(dataGridViewStudentPoint, "select_student_point_all");
            }
        }

        

        

        

        

        
       
        




















    }
}
