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
    public partial class FormClassSchedule : Form
    {
        private Common _common;
        private AppMain _appMain;

        #region Property
        private string sClassEmployeeCPNO;
        private string sClassEmployeeCLNO;

        public string ClassEmployeeCPNO
        {
            get { return sClassEmployeeCPNO; }
            set { sClassEmployeeCPNO = value; }
        }
        public string ClassEmployeeCLNO
        {
            get { return sClassEmployeeCLNO; }
            set { sClassEmployeeCLNO = value; }
        }

        #endregion

        public FormClassSchedule()
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
                case "dataGridViewClassStudy":
                    dataGridViewClassSchedule.Rows.Clear();                    
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
            string businessCD = comboBoxCampusType.SelectedValue.ToString();
            string cpno = comboBoxCampus.SelectedValue.ToString();

            switch (pQueryKind)
            {
 
                case "select_class_study":

                    //반 차시 정보 조회(과정1) 
                    pSqlCommand.CommandText = @"                       
		                SELECT (SELECT usernm FROM tls_member WHERE userid = CS.tid) AS TID
		                     , (SELECT cpnm FROM tls_campus WHERE cpno = CS.cpno) AS CPNM
                             , CS.term_cd
			                 , TC.clnm
			                 , STUFF(STUFF(CS.sdate, 5, 0, '-'), 8, 0, '-') AS SDATE
			                 , STUFF(STUFF(CS.edate, 5, 0, '-'), 8, 0, '-') AS EDATE
			                 , DBO.F_U_WEEK_HAN(CS.week_day) AS WEEK_DAY
			                 , (TS.sdnm + view_sdnm) AS SDNM
                             , CS.sdno
			                 , CS.j_use_yn
			                 , CS.j_count
			                 , CS.j_hitpoint
			                 , CS.j_quiz_cnt
			                 , CS.correct_yn
			                 , CS.c_use_yn
			                 , CS.c_common_cnt
			                 , CS.c_each_cnt
			                 , CS.l_quiz_cnt
			                 , CS.concept_yn
			                 , CS.quiz_yn
			                 , CS.menu_yn
			                 , (SELECT usernm FROM tls_member WHERE userid = CS.rid) AS RID
			                 , CS.RDATETIME
			                 , (SELECT usernm FROM tls_member WHERE userid = CS.uid) AS UID
			                 , CS.UDATETIME
                             , CS.yyyy
                             , CS.cpno
                             , CS.clno
		                  FROM tls_class_study AS CS
                     LEFT JOIN tls_class AS TC
	                        ON CS.cpno = TC.cpno and CS.clno = TC.clno
	                 LEFT JOIN tls_study AS TS
	                        ON CS.sdno = TS.sdno
		                 WHERE CS.cpno = " + ClassEmployeeCPNO + @"                    
                           AND CS.clno = " + ClassEmployeeCLNO + @"
                           AND CONVERT(CHAR,GETDATE(), 112) BETWEEN CS.sdate AND CS.edate		            
                        ORDER BY TC.clnm, CS.sdate
                    ";
                    break;

               case "select_class_study_all":

                    //반별, 학습별 차시 정보 조회(과정1) 
                    pSqlCommand.CommandText = @"                       
		                SELECT (SELECT usernm FROM tls_member WHERE userid = CS.tid) AS TID
		                     , (SELECT cpnm FROM tls_campus WHERE cpno = CS.cpno) AS CPNM
                             , CS.term_cd
			                 , TC.clnm
			                 , STUFF(STUFF(CS.sdate, 5, 0, '-'), 8, 0, '-') AS SDATE
			                 , STUFF(STUFF(CS.edate, 5, 0, '-'), 8, 0, '-') AS EDATE
			                 , DBO.F_U_WEEK_HAN(CS.week_day) AS WEEK_DAY
			                 , (TS.sdnm + '-' + view_sdnm) AS SDNM
                             , CS.sdno
			                 , CS.j_use_yn
			                 , CS.j_count
			                 , CS.j_hitpoint
			                 , CS.j_quiz_cnt
			                 , CS.correct_yn
			                 , CS.c_use_yn
			                 , CS.c_common_cnt
			                 , CS.c_each_cnt
			                 , CS.l_quiz_cnt
			                 , CS.concept_yn
			                 , CS.quiz_yn
			                 , CS.menu_yn
			                 , (SELECT usernm FROM tls_member WHERE userid = CS.rid) AS rid
			                 , CS.RDATETIME
			                 , (SELECT usernm FROM tls_member WHERE userid = CS.uid) AS uid
			                 , CS.UDATETIME
                             , CS.yyyy
                             , CS.cpno
                             , CS.clno
		                  FROM tls_class_study AS CS
                     LEFT JOIN tls_class AS TC
	                        ON CS.cpno = TC.cpno and CS.clno = TC.clno
	                 LEFT JOIN tls_study AS TS
	                        ON CS.sdno = TS.sdno
		                 WHERE CS.cpno = " + ClassEmployeeCPNO + @"
                    ";                    
                    if (!string.IsNullOrEmpty(textBoxClassNM.Text))
                    {
                        pSqlCommand.CommandText += @"
                         AND TC.clnm LIKE '%" + textBoxClassNM.Text + "%' ";
                    }
                    if (!string.IsNullOrEmpty(textBoxStudyNM.Text))
                    {
                        pSqlCommand.CommandText += @"
                         AND sdnm LIKE '%" + textBoxStudyNM.Text + "%' ";
                    }                    
                    pSqlCommand.CommandText += @"                      
                           AND REPLACE(CONVERT(VARCHAR(10), '" + dateTimePickerClassStudy.Value + @"', 112), '-', '') BETWEEN CS.sdate AND CS.edate		            
                        ORDER BY TC.clnm, CS.sdate
                    ";
                    textBoxClassNM.Text = "";                    
                    textBoxStudyNM.Text = "";
                    break;

               case "select_class_schedule":

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
		                  FROM tls_class_schedule AS A
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
	                     WHERE A.yyyy = '" + GetCellValue(dataGridViewClassStudy, dataGridViewClassStudy.CurrentCell.RowIndex, "yyyy") + @"'
	                       AND A.term_cd = '" + GetCellValue(dataGridViewClassStudy, dataGridViewClassStudy.CurrentCell.RowIndex, "term_cd") + @"'
		                   AND A.cpno = '" + GetCellValue(dataGridViewClassStudy, dataGridViewClassStudy.CurrentCell.RowIndex, "cpno") + @"'
		                   AND A.clno = '" + GetCellValue(dataGridViewClassStudy, dataGridViewClassStudy.CurrentCell.RowIndex, "clno") + @"'
		                   AND A.sdno = '" + GetCellValue(dataGridViewClassStudy, dataGridViewClassStudy.CurrentCell.RowIndex, "sdno") + @"'
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
        private void FormClassSchedule_Load(object sender, EventArgs e)
        {
            InitCombo();
            //반 차시 조회
            SelectDataGridView(dataGridViewClassStudy, "select_class_study");
        }

        private void textBoxClassNM_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SelectDataGridView(dataGridViewClassStudy, "select_class_study_all");
            }
        }

        private void textBoxStudyNM_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SelectDataGridView(dataGridViewClassStudy, "select_class_study_all");
            }
        }

        private void buttonClassStudy_Click(object sender, EventArgs e)
        {
            SelectDataGridView(dataGridViewClassStudy, "select_class_study_all");
        }

        private void dataGridViewClassStudy_Click(object sender, EventArgs e)
        {
            SelectDataGridView(dataGridViewClassSchedule, "select_class_schedule");
        }
        


        #endregion Method

        
















    }
}
