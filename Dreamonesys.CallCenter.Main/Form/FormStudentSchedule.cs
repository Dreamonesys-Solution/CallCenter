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
    public partial class FormStudentSchedule : Form
    {
        private Common _common;
        private AppMain _appMain;

        public string StudyType { get; set; }

        #region Property
        private string sClassStudentCPNO;        
        private string sClassStudentUID;

        public string ClassStudentCPNO
        {
            get { return sClassStudentCPNO; }
            set { sClassStudentCPNO = value; }
        }
        
        public string ClassStudentUID
        {
            get { return sClassStudentUID; }
            set { sClassStudentUID = value; }
        }
       
        #endregion

        public FormStudentSchedule()
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


        #region Method

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
                case "dataGridViewStudentStudy":
                    dataGridViewStudentSchedule.Rows.Clear();
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
                                //dataGridViewCampus.Rows[dataGridViewCampus.Rows.Count - 1].Cells[colCount].Value = 
                                row[pDataGridView.Columns[colCount].DataPropertyName].ToString();
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

            switch (pQueryKind)
            {
 
                case "select_student_study":

                    //학생 차시 정보 조회(과정2)
                    pSqlCommand.CommandText = @"                       
		                SELECT (SELECT usernm FROM tls_member WHERE userid = MS.tid) AS TID
	    	                 , (SELECT cpnm FROM tls_campus WHERE cpno = MS.cpno) AS CPNM
                             , MS.term_cd
		                     , TC.clnm
                             , (SELECT usernm FROM tls_member where userid = ms.userid) AS USERNM
			                 , STUFF(STUFF(MS.sdate, 5, 0, '-'), 8, 0, '-') AS SDATE 
	                         , STUFF(STUFF(MS.edate, 5, 0, '-'), 8, 0, '-') AS EDATE
	 		                 , DBO.F_U_WEEK_HAN(MS.week_day) AS WEEK_DAY
			                 , (TS.sdnm + view_sdnm) AS SDNM
                             , MS.sdno
			                 , MS.use_yn
			                 , MS.j_use_yn
			                 , MS.j_count
			                 , MS.j_hitpoint
			                 , MS.j_quiz_cnt
			                 , MS.correct_yn
			                 , MS.m_use_yn
			                 , MS.m_count			 
			                 , MS.m_hitpoint
			                 , MS.m_quiz_cnt
			                 , MS.m_quiz_type
			                 , MS.l_quiz_cnt
			                 , MS.concept_yn
			                 , MS.quiz_yn
			                 , MS.menu_yn
			                 , (SELECT usernm FROM tls_member WHERE userid = MS.rid) AS RID
			                 , MS.RDATETIME
			                 , (SELECT usernm FROM tls_member WHERE userid = MS.uid) AS UID
			                 , MS.UDATETIME
                             , MS.yyyy
                             , MS.cpno
                             , MS.userid
	                     FROM tls_member_study AS MS
                    LEFT JOIN tls_class AS TC
	                       ON MS.cpno = TC.cpno and MS.clno = TC.clno
	                LEFT JOIN tls_study AS TS
	                       ON MS.sdno = TS.sdno
		                WHERE MS.cpno = " + ClassStudentCPNO + @"
                          AND MS.userid = " + ClassStudentUID + @"
		                  AND CONVERT(CHAR, GETDATE(), 112) BETWEEN MS.sdate AND MS.edate
                        ORDER BY MS.sdate
		            ";
                    break;

                case "select_student_study_all":

                    //반별 학생, 학습명 차시 정보 조회(과정2)
                    pSqlCommand.CommandText = @"                       
		                SELECT (SELECT usernm FROM tls_member WHERE userid = MS.tid) AS TID
	    	                 , (SELECT cpnm FROM tls_campus WHERE cpno = MS.cpno) AS CPNM
                             , MS.term_cd
		                     , TC.clnm
                             , TM.usernm
			                 , STUFF(STUFF(MS.sdate, 5, 0, '-'), 8, 0, '-') AS SDATE 
	                         , STUFF(STUFF(MS.edate, 5, 0, '-'), 8, 0, '-') AS EDATE
	 		                 , DBO.F_U_WEEK_HAN(MS.week_day) AS WEEK_DAY
			                 , (TS.sdnm + '-' + view_sdnm) AS SDNM
                             , MS.sdno
			                 , MS.use_yn
			                 , MS.j_use_yn
			                 , MS.j_count
			                 , MS.j_hitpoint
			                 , MS.j_quiz_cnt
			                 , MS.correct_yn
			                 , MS.m_use_yn
			                 , MS.m_count			 
			                 , MS.m_hitpoint
			                 , MS.m_quiz_cnt
			                 , MS.m_quiz_type
			                 , MS.l_quiz_cnt
			                 , MS.concept_yn
			                 , MS.quiz_yn
			                 , MS.menu_yn
			                 , (SELECT usernm FROM tls_member WHERE userid = MS.rid) AS RID
			                 , MS.RDATETIME
			                 , (SELECT usernm FROM tls_member WHERE userid = MS.uid) AS UID
			                 , MS.UDATETIME
                             , MS.yyyy
                             , MS.cpno
                             , MS.userid
	                     FROM tls_member_study AS MS
                    LEFT JOIN tls_member AS TM
                           ON MS.userid = TM.userid
                    LEFT JOIN tls_class AS TC
	                       ON MS.cpno = TC.cpno and MS.clno = TC.clno
	                LEFT JOIN tls_study AS TS
	                       ON MS.sdno = TS.sdno
		                WHERE MS.cpno = " + ClassStudentCPNO + @"
                    ";
                    if (!string.IsNullOrEmpty(textBoxClassNM.Text))
                    {
                        pSqlCommand.CommandText += @"
                         AND TC.clnm LIKE '%" + textBoxClassNM.Text + "%' ";
                    }
                    if (!string.IsNullOrEmpty(textBoxStudentNM.Text))
                    {
                        pSqlCommand.CommandText += @"
                         AND tm.usernm LIKE '%" + textBoxStudentNM.Text + "%' ";
                    }
                    if (!string.IsNullOrEmpty(textBoxStudyNM.Text))
                    {
                        pSqlCommand.CommandText += @"
                         AND sdnm LIKE '%" + textBoxStudyNM.Text + "%' ";
                    }
                    pSqlCommand.CommandText += @"                      
                           AND REPLACE(CONVERT(VARCHAR(10), '" + dateTimePickerStudentStudy.Value + @"', 112), '-', '') BETWEEN MS.sdate AND MS.edate		            
                        ORDER BY TC.clnm, usernm, MS.sdate
		            ";
                    textBoxClassNM.Text = "";
                    textBoxStudentNM.Text = "";
                    textBoxStudyNM.Text = "";
                    break;

                case "select_student_schedule":

                    //반 차시 리스트 
                    pSqlCommand.CommandText += @"
                        SELECT STUFF(STUFF(A.cdate, 5, 0, '-'), 8, 0, '-') AS CDATE
			                 , DBO.F_U_WEEK_HAN(A.cweek_cd) AS CWEEK_CD
			                 , CONVERT(VARCHAR(2), F.sort) + '단원' + ' ' +  CONVERT(VARCHAR(2), E.time_cnt) + '차시' AS TIME_CNT
			                 , B.bknm
			                 , D.sdnm + '-' + D.view_sdnm AS SDNM
			                 , (SELECT name1 FROM tbl_quiz_class WHERE classa = F.classa) AS CLASSA
			                 , G.view_unnm
			                 , H.view_dfnm AS DFNM
			                 , CASE G.room_code WHEN 'P' THEN 'SMART Room'
					 	                        WHEN 'I' THEN 'LESSON Room'
						                        WHEN 'S' THEN 'SDL Room'
			                   END AS ROOM_CODE
			                 , CASE G.study_type_1 WHEN 'PC' THEN 'PC-(동+문)'
						                           WHEN 'PM' THEN 'PM-(동+문)'
						                           WHEN 'PQ' THEN 'PQ-(동+문)'
						                           WHEN 'E' THEN '채점'
							                       WHEN 'R' THEN '단말기'
							                       WHEN 'Q' THEN '문제풀이'
							                       WHEN 'M' THEN '동영상'
			                   END AS study_type_1
			                 , CASE G.study_type_2 WHEN 'PC' THEN 'PC-(동+문)'
							                       WHEN 'PM' THEN 'PM-(동+문)'
							                       WHEN 'PQ' THEN 'PQ-(동+문)'
							                       WHEN 'E' THEN '채점'
							                       WHEN 'R' THEN '단말기'
							                       WHEN 'Q' THEN '문제풀이'
							                       WHEN 'M' THEN '동영상'
			                   END AS study_type_2			                         		
		                  FROM tls_member_schedule AS A
                    INNER JOIN tls_book AS B 
	                        ON B.bkno = A.bkno
                    INNER JOIN tls_study AS D 
                            ON D.sdno = A.sdno
                    INNER JOIN tls_schedule AS E 
		                    ON E.sdno = A.sdno AND E.scno = A.scno
                    INNER JOIN TLS_CHAPTER AS F 
		                    ON F.bkno = B.bkno AND F.chno = E.chno
                    INNER JOIN tls_unit AS G 
		                    ON G.bkno = B.bkno AND G.chno = F.chno AND G.unno = E.unno
                    INNER JOIN tls_lvl_def AS H 
		                    ON H.lvno = G.lvno AND H.dfno = G.dfno
	                     WHERE A.yyyy = '" + GetCellValue(dataGridViewStudentStudy, dataGridViewStudentStudy.CurrentCell.RowIndex, "yyyy") + @"'
	                       AND A.term_cd = '" + GetCellValue(dataGridViewStudentStudy, dataGridViewStudentStudy.CurrentCell.RowIndex, "term_cd") + @"'
		                   AND A.cpno = '" + GetCellValue(dataGridViewStudentStudy, dataGridViewStudentStudy.CurrentCell.RowIndex, "cpno") + @"'
		                   AND A.userid = '" + GetCellValue(dataGridViewStudentStudy, dataGridViewStudentStudy.CurrentCell.RowIndex, "userid") + @"'
		                   AND A.sdno = '" + GetCellValue(dataGridViewStudentStudy, dataGridViewStudentStudy.CurrentCell.RowIndex, "sdno") + @"'
	                     ORDER BY cdate, G.sort

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

        #region Event

        /// <summary>
        /// 폼 로드
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <history>
        /// 박석제, 2014-09-24, 생성
        /// </history>
        private void FormStudentSchedule_Load(object sender, EventArgs e)
        {
            //반 차시 조회
            switch (StudyType)
            {
                case "C": //반 차시 조회
                    tabControl1.SelectedTab = tabPageStudentSchedule;
                    SelectDataGridView(dataGridViewStudentStudy, "select_class_study");
                    break;
                case "S": //학생 차시 조회
                    tabControl1.SelectedTab = tabPage2;
                    SelectDataGridView(dataGridViewStudentStudy, "select_student_study");
                    break;
                default:
                    break;
            }
        }

        private void dataGridViewStudentStudy_Click(object sender, EventArgs e)
        {
            //학생 차시 리스트 조회
            SelectDataGridView(dataGridViewStudentSchedule, "select_student_schedule");
        }

        private void textBoxClassNM_KeyDown(object sender, KeyEventArgs e)
        {   //반별 차시 조회
            if (e.KeyCode == Keys.Enter)
            {
                SelectDataGridView(dataGridViewStudentStudy, "select_student_study_all");
            }            
        }

        private void textBoxStudentNM_KeyDown(object sender, KeyEventArgs e)
        {   //학생별 차시 조회
            if (e.KeyCode == Keys.Enter)
            {
                SelectDataGridView(dataGridViewStudentStudy, "select_student_study_all");
            }            
        }

        private void textBoxStudyNM_KeyDown(object sender, KeyEventArgs e)
        {   //학습별 차시 조회
            if (e.KeyCode == Keys.Enter)
            {

                SelectDataGridView(dataGridViewStudentStudy, "select_student_study_all");
            }            
        }

        private void buttonStudentStudy_Click(object sender, EventArgs e)
        {   //날짜별 차시 조회
            SelectDataGridView(dataGridViewStudentStudy, "select_student_study_all");
        }

        

        #endregion Method

        

        







    }
}