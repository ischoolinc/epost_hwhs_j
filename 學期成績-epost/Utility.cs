﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;
using K12.Data;
using System.IO;
using System.Xml.Linq;
using JHSchool.Data;

namespace hwhs_epost_semester
{
    public class Utility
    {
        
        /// <summary>
        /// 透過學生編號、開始與結束日期，取得學習服務統計值
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, decimal> GetServiceLearningDetailByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, decimal> retVal = new Dictionary<string, decimal>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,occur_date,reason,hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') and occur_date >='" + beginDate.ToShortDateString() + "' and occur_date <='" + endDate.ToShortDateString() + "'order by ref_student_id,occur_date;";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    decimal hr;

                    string sid = dr[0].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, 0);

                    if (decimal.TryParse(dr["hours"].ToString(), out hr))
                        retVal[sid] += hr;
                }
            }
            return retVal;
        }


        /// <summary>
        /// 透過日期區間取得獎懲資料,傳入學生ID,開始日期,結束日期,回傳：學生ID,獎懲統計名稱,統計值
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetDisciplineCountByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<string> nameList = new string[] { "大功", "小功", "嘉獎", "大過", "小過", "警告", "留校" }.ToList();

            // 取得獎懲資料
            List<DisciplineRecord> dataList = Discipline.SelectByStudentIDs(StudentIDList);

            foreach (DisciplineRecord data in dataList)
            {
                if (data.OccurDate >= beginDate && data.OccurDate <= endDate)
                {
                    // 初始化
                    if (!retVal.ContainsKey(data.RefStudentID))
                    {
                        retVal.Add(data.RefStudentID, new Dictionary<string, int>());
                        foreach (string str in nameList)
                            retVal[data.RefStudentID].Add(str, 0);
                    }

                    // 獎勵
                    if (data.MeritFlag == "1")
                    {
                        if (data.MeritA.HasValue)
                            retVal[data.RefStudentID]["大功"] += data.MeritA.Value;

                        if (data.MeritB.HasValue)
                            retVal[data.RefStudentID]["小功"] += data.MeritB.Value;

                        if (data.MeritC.HasValue)
                            retVal[data.RefStudentID]["嘉獎"] += data.MeritC.Value;
                    }
                    else if (data.MeritFlag == "0")
                    { // 懲戒
                        if (data.Cleared != "是")
                        {
                            if (data.DemeritA.HasValue)
                                retVal[data.RefStudentID]["大過"] += data.DemeritA.Value;

                            if (data.DemeritB.HasValue)
                                retVal[data.RefStudentID]["小過"] += data.DemeritB.Value;

                            if (data.DemeritC.HasValue)
                                retVal[data.RefStudentID]["警告"] += data.DemeritC.Value;
                        }
                    }
                    else if (data.MeritFlag == "2")
                    {
                        // 留校察看
                        retVal[data.RefStudentID]["留校"]++;
                    }
                }
            }
            return retVal;
        }

        /// <summary>
        /// 透過日期區間取得學生缺曠統計(傳入學生系統編號、開始日期、結束日期；回傳：學生系統編號、獎懲名稱,統計值
        /// </summary>
        /// <param name="StudIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetAttendanceCountByDate(List<StudentRecord> StudRecordList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<PeriodMappingInfo> PeriodMappingList = PeriodMapping.SelectAll();
            // 節次>類別
            Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
            foreach (PeriodMappingInfo rec in PeriodMappingList)
            {
                if (!PeriodMappingDict.ContainsKey(rec.Name))
                    PeriodMappingDict.Add(rec.Name, rec.Type);
            }

            List<AttendanceRecord> attendList = K12.Data.Attendance.SelectByDate(StudRecordList, beginDate, endDate);

            // 計算統計資料
            foreach (AttendanceRecord rec in attendList)
            {
                if (!retVal.ContainsKey(rec.RefStudentID))
                    retVal.Add(rec.RefStudentID, new Dictionary<string, int>());

                foreach (AttendancePeriod per in rec.PeriodDetail)
                {
                    if (!PeriodMappingDict.ContainsKey(per.Period))
                        continue;

                    // ex.一般:曠課
                    //string key = "區間" + PeriodMappingDict[per.Period] + "_" + per.AbsenceType;

                    string key = PeriodMappingDict[per.Period] + per.AbsenceType;
                    if (!retVal[rec.RefStudentID].ContainsKey(key))
                        retVal[rec.RefStudentID].Add(key, 0);

                    retVal[rec.RefStudentID][key]++;
                }
            }

            return retVal;
        }

        /// <summary>
        /// 透過 ClassID 取得學生為一般，依照年級、班級名稱、座號排序
        /// </summary>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static List<string> GetClassStudentIDList1ByClassID(List<string> ClassIDList)
        {
            List<string> retVal = new List<string>();
            QueryHelper qh = new QueryHelper();
            string query = "select student.id from student inner join class on student.ref_class_id = class.id where student.status=1 and class.id in(" + string.Join(",", ClassIDList.ToArray()) + ") order by class.grade_year,class.class_name,student.seat_no";
            DataTable dt = new DataTable();
            dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
                retVal.Add(dr[0].ToString());

            dt.Clear();
            return retVal;
        }


        /// <summary>
        /// 取得系統內學生類別,群組用[]表示,沒有群組直接名稱
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, List<string>> GetStudentTagRefDict()
        {
            // 學生類別,StudentID
            Dictionary<string, List<string>> retVal = new Dictionary<string, List<string>>();
            QueryHelper qh = new QueryHelper();
            string query = "select tag.prefix,tag.name,ref_student_id from tag left join tag_student on tag.id = tag_student.ref_tag_id order by tag.prefix,tag.name";
            DataTable dt = new DataTable();
            dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
            {
                string strP = "", key = "", StudID = "";

                if (dr["prefix"] != null)
                    strP = dr["prefix"].ToString();

                if (string.IsNullOrEmpty(strP))
                    key = dr["name"].ToString();
                else
                    key = "[" + strP + "]";

                if (dr["ref_student_id"] != null)
                    StudID = dr["ref_student_id"].ToString();

                if (!retVal.ContainsKey(key))
                    retVal.Add(key, new List<string>());

                if (!string.IsNullOrEmpty(StudID))
                    retVal[key].Add(StudID);

            }
            return retVal;
        }

        public static Dictionary<string, Dictionary<string, DAO.SubjectDomainName>> GetStudentSCAttendCourse(List<string> StudentIDList, List<string> CourseIDList, string examID)
        {
            Dictionary<string, Dictionary<string, DAO.SubjectDomainName>> retVal = new Dictionary<string, Dictionary<string, DAO.SubjectDomainName>>();
            QueryHelper qh = new QueryHelper();
            string query = "select ref_student_id,course.domain,course.subject,course.credit from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join te_include on course.ref_exam_template_id = te_include.ref_exam_template_id where sc_attend.ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and course.id in(" + string.Join(",", CourseIDList.ToArray()) + ") and te_include.ref_exam_id=" + examID;
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                string id = dr[0].ToString();

                if (!retVal.ContainsKey(id))
                    retVal.Add(id, new Dictionary<string, DAO.SubjectDomainName>());

                string domainName = dr["domain"].ToString();
                string subjectName = dr["subject"].ToString();

                if (string.IsNullOrEmpty(domainName))
                    domainName = "彈性課程";



                if (!retVal[id].ContainsKey(subjectName))
                {
                    DAO.SubjectDomainName sdn = new DAO.SubjectDomainName();
                    sdn.SubjectName = subjectName;
                    sdn.DomainName = domainName;
                    decimal credit;
                    if (decimal.TryParse(dr["credit"].ToString(), out credit))
                    {
                        sdn.Credit = credit;
                    }

                    retVal[id].Add(subjectName, sdn);
                }
            }

            return retVal;
        }

        /// <summary>
        /// 透過學生系統編號、學年度、學期，取得考試的科目名稱
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, List<string>> GetExamSubjecList(List<string> StudentIDList, int SchoolYear, int Semester)
        {
            Dictionary<string, List<string>> retVal = new Dictionary<string, List<string>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select distinct ref_exam_id,course.subject from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join te_include on course.ref_exam_template_id = te_include.ref_exam_template_id where sc_attend.ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and course.school_year=" + SchoolYear + " and  course.semester=" + Semester + " and course.subject is not null order by ref_exam_id,subject";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr[0].ToString();

                    if (!retVal.ContainsKey(id))
                        retVal.Add(id, new List<string>());

                    string subjectName = dr["subject"].ToString();

                    if (!retVal[id].Contains(subjectName))
                        retVal[id].Add(subjectName);
                }
            }
            return retVal;
        }

        /// <summary>
        /// 直接傳遞弘文國中學期成績單EPOST 需要的科目名稱
        /// </summary>       
        /// <returns></returns>
        public static List<string> GetSemesterSubjecList()
        {
            List<string> SemesterSubjecList = new List<string>();

            SemesterSubjecList.Add("國文");
            SemesterSubjecList.Add("英語");
            SemesterSubjecList.Add("數學");
            SemesterSubjecList.Add("社會");
            SemesterSubjecList.Add("自然科學");
            SemesterSubjecList.Add("理化");
            SemesterSubjecList.Add("自然");
            SemesterSubjecList.Add("資訊科技");
            SemesterSubjecList.Add("生活科技");
            SemesterSubjecList.Add("音樂");
            SemesterSubjecList.Add("視覺藝術");
            SemesterSubjecList.Add("表演藝術");
            SemesterSubjecList.Add("家政");
            SemesterSubjecList.Add("童軍");
            SemesterSubjecList.Add("輔導");
            SemesterSubjecList.Add("健康教育");
            SemesterSubjecList.Add("體育");
            SemesterSubjecList.Add("英語聽講");
            SemesterSubjecList.Add("資訊應用");
            SemesterSubjecList.Add("ESL");
            SemesterSubjecList.Add("地球科學");
            SemesterSubjecList.Add("閱讀理解");
            SemesterSubjecList.Add("閱讀與寫作");
            SemesterSubjecList.Add("語文表達");

            return SemesterSubjecList;
        }
        
        /// <summary>
        /// 取得學生學期科目名稱和排序
        /// </summary>
        /// <param name="allStudentID"></param>
        /// <param name="schoolYear"></param>
        /// <param name="semester"></param>
        /// <returns></returns>
        public static List<string> GetFormSubjectList(List<string> allStudentID, int schoolYear, int semester)
        {
            List<string> SemesterSubjecList = new List<string>();
            List<JHSemesterScoreRecord> semesterScoreList = JHSemesterScore.SelectBySchoolYearAndSemester(allStudentID, schoolYear, semester);

            foreach (JHSemesterScoreRecord record in semesterScoreList)
            {
                foreach (var item in record.Subjects)
                {
                    if (!SemesterSubjecList.Contains(item.Key))
                        SemesterSubjecList.Add(item.Key);
                }
            }

            SemesterSubjecList.Sort(new StringComparer(GetSubjectOrder().ToArray()));

            return SemesterSubjecList;
        }

        /// <summary>
        /// 科目排序
        /// </summary>
        /// <returns></returns>
        public static List<string> GetSubjectOrder()
        {
            List<string> result = new List<string>();
            QueryHelper qh = new QueryHelper();
            string sql =
                @"
WITH subject_mapping AS 
(
SELECT
    unnest(xpath('//Subjects/Subject/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS subject_name
FROM  
    list 
WHERE name  ='JHEvaluation_Subject_Ordinal'
)SELECT
		replace (subject_name ,'&amp;amp;','&') AS subject_name
	FROM  subject_mapping
";

            DataTable dt = qh.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                string subject = dr["subject_name"].ToString();
                result.Add(subject);
            }
            return result;
        }



        /// <summary>
        /// 取得評量比例設定
        /// </summary>
        public static Dictionary<string, decimal> GetScorePercentageHS()
        {
            Dictionary<string, decimal> returnData = new Dictionary<string, decimal>();
            FISCA.Data.QueryHelper qh1 = new FISCA.Data.QueryHelper();
            string query1 = @"select id,CAST(regexp_replace( xpath_string(exam_template.extension,'/Extension/ScorePercentage'), '^$', '0') as integer) as ScorePercentage  from exam_template";
            System.Data.DataTable dt1 = qh1.Select(query1);

            foreach (System.Data.DataRow dr in dt1.Rows)
            {
                string id = dr["id"].ToString();
                decimal sp = 50;
                if (decimal.TryParse(dr["ScorePercentage"].ToString(), out sp))
                    returnData.Add(id, sp);
                else
                    returnData.Add(id, 50);

            }
            return returnData;
        }

        /// <summary>
        /// 取得等第對照
        /// </summary>
        /// <returns></returns>
        public static Dictionary<decimal, string> GetScoreLevelMapping()
        {
            Dictionary<decimal, string> retVal = new Dictionary<decimal, string>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select content from list where name='等第對照表';";
            DataTable dt = qh.Select(strSQL);
            string xmlStr = "";
            foreach (DataRow dr in dt.Rows)
                xmlStr = dr[0].ToString();

            try
            {
                if (xmlStr != "")
                {
                    XElement elmRoot = XElement.Parse(xmlStr);
                    if (elmRoot != null)
                    {
                        string elmstr1 = elmRoot.Element("Configuration").Value;
                        XElement elmXml = XElement.Parse(elmstr1);
                        Dictionary<decimal, string> tmp = new Dictionary<decimal, string>();
                        foreach (XElement elm in elmXml.Elements("ScoreMapping"))
                        {
                            decimal dd;
                            if (elm.Attribute("Score") != null && elm.Attribute("Name") != null)
                                if (decimal.TryParse(elm.Attribute("Score").Value, out dd))
                                    tmp.Add(dd, elm.Attribute("Name").Value);
                        }
                        List<decimal> tmpsort = new List<decimal>();

                        tmpsort = (from data in tmp orderby data.Key descending select data.Key).ToList();

                        foreach (decimal d in tmpsort)
                        {
                            if (tmp.ContainsKey(d))
                                retVal.Add(d, tmp[d]);
                        }

                        if (!retVal.ContainsKey(0))
                            retVal.Add(0, "丁");
                    }

                }
            }
            catch { }
            return retVal;

        }


    }
}
