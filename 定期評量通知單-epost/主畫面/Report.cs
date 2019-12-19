using System;
using System.Data;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Aspose.Words;
using FISCA.DSAUtil;
using K12.Data;
using K12.Data.Configuration;
using SmartSchool.ePaper;
using FISCA.Data;
using JHSchool.Data;
using HsinChu.JHEvaluation.Data;
using System.Linq;
using JHSchool.Evaluation.Calculation;

namespace hwhs.epost.定期評量通知單
{
    internal class Report : IReport
    {
        private BackgroundWorker _BGWAbsenceNotification;

        private List<StudentRecord> SelectedStudents { get; set; }

        private Dictionary<string, List<string>> config;

        private static QueryHelper queryHelper;

        string addconfigName = "定期評量通知單_缺曠別設定_2019_弘文epost";

        List<string> configkeylist { get; set; }
        ConfigOBJ obj;

        string entityName;

        //轉縮寫
        Dictionary<string, string> absenceList = new Dictionary<string, string>();

        public Report(string _entityName)
        {
            entityName = _entityName;
        }

        public void Print(List<string> StudIDList)
        {
            #region IReport 成員

            AbsenceNotificationSelectDateRangeForm form = new AbsenceNotificationSelectDateRangeForm(StudIDList);

            if (form.ShowDialog() == DialogResult.OK)
            {
                queryHelper = new QueryHelper();
                #region 讀取缺曠別 Preference
                config = new Dictionary<string, List<string>>();
                configkeylist = new List<string>();
                //XmlElement preferenceData = CurrentUser.Instance.Preference["缺曠通知單_缺曠別設定"];
                ConfigData cd = K12.Data.School.Configuration[addconfigName];
                XmlElement preferenceData = cd.GetXml("XmlData", null);

                if (preferenceData != null)
                {
                    foreach (XmlElement type in preferenceData.SelectNodes("Type"))
                    {
                        string prefix = type.GetAttribute("Text");
                        if (!config.ContainsKey(prefix))
                        {
                            configkeylist.Add(prefix);
                            config.Add(prefix, new List<string>());
                        }

                        foreach (XmlElement absence in type.SelectNodes("Absence"))
                        {
                            if (!config[prefix].Contains(absence.GetAttribute("Text")))
                                config[prefix].Add(absence.GetAttribute("Text"));
                        }
                    }
                }
                configkeylist.Sort();
                #endregion

                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化定期評量通知單CSV檔...");





                #region 建立設定檔
                obj = new ConfigOBJ();

                obj.SchoolYear = "" + form._SelSchoolYear;
                obj.Semester = "" + form._SelSemester;
                obj.ExamName = "" + form._SelExamName;
                obj.ExamID = "" + form._SelExamID;
                obj.SelSubjNameList = form._SelSubjNameList; // 使用者勾選科目

                obj.StartDate = form.StartDate;
                obj.EndDate = form.EndDate;                
                obj.PrintHasRecordOnly = form.PrintHasRecordOnly;
                
                obj.Template = form.Template;

                obj.ReceiveName = form.ReceiveName;
                obj.ReceiveAddress = form.ReceiveAddress !=""? form.ReceiveAddress: "聯絡地址";
                obj.ConditionName = form.ConditionName;
                obj.ConditionNumber = form.ConditionNumber;
                obj.ConditionName2 = form.ConditionName2;
                obj.ConditionNumber2 = form.ConditionNumber2;
                obj.PrintStudentList = form.PrintStudentList;
                #endregion

                _BGWAbsenceNotification = new BackgroundWorker();
                _BGWAbsenceNotification.DoWork += new DoWorkEventHandler(_BGWAbsenceNotification_DoWork);
                _BGWAbsenceNotification.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CommonMethods.WordReport_RunWorkerCompleted);
                _BGWAbsenceNotification.ProgressChanged += new ProgressChangedEventHandler(CommonMethods.Report_ProgressChanged);
                _BGWAbsenceNotification.WorkerReportsProgress = true;
                _BGWAbsenceNotification.RunWorkerAsync();
            }

            #endregion
        }

        private void _BGWAbsenceNotification_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 取得學生

            if (entityName.ToLower() == "student") //學生模式
            {
                SelectedStudents = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
            }
            else if (entityName.ToLower() == "class") //班級模式
            {
                SelectedStudents = new List<StudentRecord>();
                foreach (StudentRecord each in Student.SelectByClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource))
                {
                    if (each.Status != StudentRecord.StudentStatus.一般)
                        continue;

                    SelectedStudents.Add(each);
                }
            }
            else
                throw new NotImplementedException();

            SelectedStudents.Sort(new Comparison<StudentRecord>(CommonMethods.ClassSeatNoComparer));
            
            _BGWAbsenceNotification.ReportProgress(10);

            #endregion


            #region 快取資料

            //超級資訊物件
            Dictionary<string, StudentOBJ> StudentSuperOBJ = new Dictionary<string, StudentOBJ>();

            // 獎懲List
            List<string> meritNameList = new List<string>();

            //合併列印的資料
            Dictionary<string, object> Allmapping = new Dictionary<string, object>();
            Dictionary<string, string> ReversionDic = new Dictionary<string, string>();

            //所有學生ID
            List<string> allStudentID = new List<string>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

            #region 取得 Period List            
            Dictionary<string, string> TestPeriodList = new Dictionary<string, string>();
            Dictionary<string, object> mappingAccessory_copy = new Dictionary<string, object>();

            foreach (K12.Data.PeriodMappingInfo each in K12.Data.PeriodMapping.SelectAll())
            {
                if (!TestPeriodList.ContainsKey(each.Name)) //節次<-->類別
                    TestPeriodList.Add(each.Name, each.Type);
            }
            #endregion

            #region 取得 Absence List
            Dictionary<string, string> TestAbsenceList = new Dictionary<string, string>(); //代碼替換(新)
            foreach (K12.Data.AbsenceMappingInfo each in K12.Data.AbsenceMapping.SelectAll())
            {
                if (!absenceList.ContainsKey(each.Name))
                {
                    absenceList.Add(each.Name, each.Abbreviation);
                }

                if (!TestAbsenceList.ContainsKey(each.Name)) //縮寫<-->假別
                {
                    TestAbsenceList.Add(each.Abbreviation, each.Name);
                }


                //Allmapping.Add("類型" + DefinedType + "缺曠" + DefinedAbsence,

            }
            #endregion

            #region 設定 Merit List (固定)
            meritNameList.Add("大功");
            meritNameList.Add("小功");
            meritNameList.Add("嘉獎");
            meritNameList.Add("大過");
            meritNameList.Add("小過");
            meritNameList.Add("警告");
            #endregion

            //????使用者所選取的所有假別種類????
            //List<string> userDefinedAbsenceList = new List<string>();

            //int DefinedType = 1;
            //foreach (string kind in configkeylist)
            //{
            //    int DefinedAbsence = 1;
            //    Allmapping.Add("類型" + DefinedType, kind);

            //    foreach (string type in config[kind])
            //    {
            //        Allmapping.Add("類型" + DefinedType + "缺曠" + DefinedAbsence, type);
            //        Allmapping.Add("類型" + DefinedType + "縮寫" + DefinedAbsence, absenceList[type]);
            //        DefinedAbsence++;

            //        if (!userDefinedAbsenceList.Contains(type))
            //        {
            //            userDefinedAbsenceList.Add(type);
            //        }
            //    }

            //    DefinedType++;
            //}

            #region 取得所有學生ID
            foreach (StudentRecord aStudent in SelectedStudents)
            {
                //建立學生資訊，班級、座號、學號、姓名、導師
                string studentID = aStudent.ID;
                if (!StudentSuperOBJ.ContainsKey(studentID))
                    StudentSuperOBJ.Add(studentID, new StudentOBJ());

                //學生ID清單
                if (!allStudentID.Contains(studentID))
                    allStudentID.Add(studentID);

                StudentSuperOBJ[studentID].student = aStudent;
                StudentSuperOBJ[studentID].TeacherName = aStudent.Class != null ? (aStudent.Class.Teacher != null ? aStudent.Class.Teacher.Name : "") : "";
                StudentSuperOBJ[studentID].ClassName = aStudent.Class != null ? aStudent.Class.Name : "";
                StudentSuperOBJ[studentID].SeatNo = aStudent.SeatNo.HasValue ? aStudent.SeatNo.Value.ToString() : "";
                StudentSuperOBJ[studentID].StudentNumber = aStudent.StudentNumber;
                StudentSuperOBJ[studentID].ParentCode = "";
            }
            #endregion

            #region 取得家長代碼
            // 因應 2019/11/14 弘文要求新epost  增加家長代碼抓取
            string ids = string.Join(",", allStudentID);

            string sql = "select student.id, student.parent_code, student.student_code, student.seat_no, student.name, class.grade_year, class.class_name from student";
            sql += " join class on class.id = student.ref_class_id where student.status in (1,2) and student.id in (" + ids + ") order by class.grade_year,class.display_order,class.class_name,student.seat_no";
            DataTable dt_parent_code = queryHelper.Select(sql); ;

            foreach (DataRow row in dt_parent_code.Rows)
            {
                if (StudentSuperOBJ.ContainsKey("" + row["id"]))
                {
                    StudentSuperOBJ["" + row["id"]].ParentCode = "" + row["parent_code"];
                }
            }

            _BGWAbsenceNotification.ReportProgress(20);
            #endregion


            #region 取得所有學生缺曠紀錄，日期區間

            List<AttendanceRecord> attendanceList = K12.Data.Attendance.SelectByDate(SelectedStudents, obj.StartDate, obj.EndDate);

            foreach (AttendanceRecord attendance in attendanceList)
            {
                if (!allStudentID.Contains(attendance.RefStudentID)) //如果是選取班級的學生
                    continue;

                string studentID = attendance.RefStudentID;
                DateTime occurDate = attendance.OccurDate;
                StudentOBJ studentOBJ = StudentSuperOBJ[studentID]; //取得這個物件

                foreach (AttendancePeriod attendancePeriod in attendance.PeriodDetail)
                {
                    string absenceType = attendancePeriod.AbsenceType; //假別
                    string periodName = attendancePeriod.Period; //節次

                    //是否為設定檔節次清單之中
                    if (!TestPeriodList.ContainsKey(periodName))
                        continue;

                    //是否為使用者選取之假別&類型
                    if (config.ContainsKey(TestPeriodList[periodName]))
                    {
                        if (config[TestPeriodList[periodName]].Contains(absenceType))
                        {
                            string PeriodAndAbsence = TestPeriodList[periodName] + "," + absenceType;
                            //區間統計
                            if (!studentOBJ.studentAbsence.ContainsKey(PeriodAndAbsence))
                            {
                                studentOBJ.studentAbsence.Add(PeriodAndAbsence, 0);
                            }

                            studentOBJ.studentAbsence[PeriodAndAbsence]++;

                            //明細記錄
                            if (!studentOBJ.studentAbsenceDetail.ContainsKey(occurDate.ToShortDateString()))
                            {
                                studentOBJ.studentAbsenceDetail.Add(occurDate.ToShortDateString(), new Dictionary<string, string>());
                            }

                            if (!studentOBJ.studentAbsenceDetail[occurDate.ToShortDateString()].ContainsKey(attendancePeriod.Period))
                            {
                                studentOBJ.studentAbsenceDetail[occurDate.ToShortDateString()].Add(attendancePeriod.Period, attendancePeriod.AbsenceType);
                            }
                        }
                    }
                }
            }

            _BGWAbsenceNotification.ReportProgress(30);

            #endregion

            List<string> DelStudent = new List<string>(); //列印的學生

            #region 無條件則全部列印
            if (obj.ConditionName == "" && obj.ConditionName2 == "")
            {
                foreach (string each1 in StudentSuperOBJ.Keys) //取出一個學生
                {
                    if (!DelStudent.Contains(each1))
                    {
                        DelStudent.Add(each1);
                    }
                }
            }
            #endregion

            #region 取得所有學生缺曠紀錄，學期累計
            foreach (AttendanceRecord attendance in K12.Data.Attendance.SelectBySchoolYearAndSemester(Student.SelectByIDs(allStudentID), int.Parse(obj.SchoolYear), int.Parse(obj.Semester)))
            {
                //1(大於),0(等於)-1(小於)
                if (obj.EndDate.CompareTo(attendance.OccurDate) == -1)
                    continue;

                string studentID = attendance.RefStudentID;
                DateTime occurDate = attendance.OccurDate;
                StudentOBJ studentOBJ = StudentSuperOBJ[studentID]; //取得這個物件

                foreach (AttendancePeriod attendancePeriod in attendance.PeriodDetail)
                {
                    string absenceType = attendancePeriod.AbsenceType; //假別
                    string periodName = attendancePeriod.Period; //節次
                    if (!TestPeriodList.ContainsKey(periodName))
                        continue;

                    string PeriodAndAbsence = TestPeriodList[periodName] + "," + absenceType;
                    //區間統計
                    if (!studentOBJ.studentSemesterAbsence.ContainsKey(PeriodAndAbsence))
                    {
                        studentOBJ.studentSemesterAbsence.Add(PeriodAndAbsence, 0);
                    }

                    studentOBJ.studentSemesterAbsence[PeriodAndAbsence]++;
                }
            }
            _BGWAbsenceNotification.ReportProgress(40);
            #endregion

            #region 取得所有學生獎懲紀錄，日期區間

            foreach (MeritRecord merit in K12.Data.Merit.SelectByOccurDate(allStudentID, obj.StartDate, obj.EndDate))
            {
                string studentID = merit.RefStudentID;
                DateTime occurDate = merit.OccurDate;
                StudentOBJ studentOBJ = StudentSuperOBJ[studentID]; //取得這個物件

                //區間統計
                if (!studentOBJ.studentMerit.ContainsKey("大功"))
                {
                    studentOBJ.studentMerit.Add("大功", 0);
                }
                if (!studentOBJ.studentMerit.ContainsKey("小功"))
                {
                    studentOBJ.studentMerit.Add("小功", 0);
                }
                if (!studentOBJ.studentMerit.ContainsKey("嘉獎"))
                {
                    studentOBJ.studentMerit.Add("嘉獎", 0);
                }

                int meritA = 0;
                int meritB = 0;
                int meritC = 0;

                studentOBJ.studentMerit["大功"] += (int)(int.TryParse("" + merit.MeritA, out meritA) ? int.Parse("" + merit.MeritA) : 0);
                studentOBJ.studentMerit["小功"] += (int)(int.TryParse("" + merit.MeritB, out meritB) ? int.Parse("" + merit.MeritB) : 0);
                studentOBJ.studentMerit["嘉獎"] += (int)(int.TryParse("" + merit.MeritC, out meritC) ? int.Parse("" + merit.MeritC) : 0);
            }

            foreach (DemeritRecord demerit in K12.Data.Demerit.SelectByOccurDate(allStudentID, obj.StartDate, obj.EndDate))
            {
                string studentID = demerit.RefStudentID;
                DateTime occurDate = demerit.OccurDate;
                StudentOBJ studentOBJ = StudentSuperOBJ[studentID]; //取得這個物件

                //區間統計
                if (!studentOBJ.studentMerit.ContainsKey("大過"))
                {
                    studentOBJ.studentMerit.Add("大過", 0);
                }
                if (!studentOBJ.studentMerit.ContainsKey("小過"))
                {
                    studentOBJ.studentMerit.Add("小過", 0);
                }
                if (!studentOBJ.studentMerit.ContainsKey("警告"))
                {
                    studentOBJ.studentMerit.Add("警告", 0);
                }

                int DemeritA = 0;
                int DemeritB = 0;
                int DemeritC = 0;

                studentOBJ.studentMerit["大過"] += (int)(int.TryParse("" + demerit.DemeritA, out DemeritA) ? int.Parse("" + demerit.DemeritA) : 0);
                studentOBJ.studentMerit["小過"] += (int)(int.TryParse("" + demerit.DemeritB, out DemeritB) ? int.Parse("" + demerit.DemeritB) : 0);
                studentOBJ.studentMerit["警告"] += (int)(int.TryParse("" + demerit.DemeritC, out DemeritC) ? int.Parse("" + demerit.DemeritC) : 0);
            }

            _BGWAbsenceNotification.ReportProgress(50);
            #endregion

            #region 取得所有學生獎懲紀錄，學期累計
            // ... 意外發現K12 API 缺曠、獎懲的支援度不同... 缺曠可以用 studentRecord 抓 ，獎懲不行
            foreach (MeritRecord merit in K12.Data.Merit.SelectBySchoolYearAndSemester(allStudentID, int.Parse(obj.SchoolYear), int.Parse(obj.Semester)))
            {
                //1(大於),0(等於)-1(小於)
                if (obj.EndDate.CompareTo(merit.OccurDate) == -1)
                    continue;

                string studentID = merit.RefStudentID;
                DateTime occurDate = merit.OccurDate;
                StudentOBJ studentOBJ = StudentSuperOBJ[studentID]; //取得這個物件

                //區間統計
                if (!studentOBJ.studentSemesterMerit.ContainsKey("大功"))
                {
                    studentOBJ.studentSemesterMerit.Add("大功", 0);
                }
                if (!studentOBJ.studentSemesterMerit.ContainsKey("小功"))
                {
                    studentOBJ.studentSemesterMerit.Add("小功", 0);
                }
                if (!studentOBJ.studentSemesterMerit.ContainsKey("嘉獎"))
                {
                    studentOBJ.studentSemesterMerit.Add("嘉獎", 0);
                }

                int meritA = 0;
                int meritB = 0;
                int meritC = 0;

                studentOBJ.studentSemesterMerit["大功"] += (int)(int.TryParse("" + merit.MeritA, out meritA) ? int.Parse("" + merit.MeritA) : 0);
                studentOBJ.studentSemesterMerit["小功"] += (int)(int.TryParse("" + merit.MeritB, out meritB) ? int.Parse("" + merit.MeritB) : 0);
                studentOBJ.studentSemesterMerit["嘉獎"] += (int)(int.TryParse("" + merit.MeritC, out meritC) ? int.Parse("" + merit.MeritC) : 0);
            }

            foreach (DemeritRecord demerit in K12.Data.Demerit.SelectBySchoolYearAndSemester(allStudentID, int.Parse(obj.SchoolYear), int.Parse(obj.Semester)))
            {
                //1(大於),0(等於)-1(小於)
                if (obj.EndDate.CompareTo(demerit.OccurDate) == -1)
                    continue;

                string studentID = demerit.RefStudentID;
                DateTime occurDate = demerit.OccurDate;
                StudentOBJ studentOBJ = StudentSuperOBJ[studentID]; //取得這個物件

                //區間統計
                if (!studentOBJ.studentSemesterMerit.ContainsKey("大過"))
                {
                    studentOBJ.studentSemesterMerit.Add("大過", 0);
                }
                if (!studentOBJ.studentSemesterMerit.ContainsKey("小過"))
                {
                    studentOBJ.studentSemesterMerit.Add("小過", 0);
                }
                if (!studentOBJ.studentSemesterMerit.ContainsKey("警告"))
                {
                    studentOBJ.studentSemesterMerit.Add("警告", 0);
                }

                int demeritA = 0;
                int demeritB = 0;
                int demeritC = 0;

                studentOBJ.studentSemesterMerit["大過"] += (int)(int.TryParse("" + demerit.DemeritA, out demeritA) ? int.Parse("" + demerit.DemeritA) : 0);
                studentOBJ.studentSemesterMerit["小過"] += (int)(int.TryParse("" + demerit.DemeritB, out demeritB) ? int.Parse("" + demerit.DemeritB) : 0);
                studentOBJ.studentSemesterMerit["警告"] += (int)(int.TryParse("" + demerit.DemeritC, out demeritC) ? int.Parse("" + demerit.DemeritC) : 0);
            }

            _BGWAbsenceNotification.ReportProgress(60);
            #endregion

            #region 取得學生通訊地址資料
            foreach (AddressRecord record in Address.SelectByStudentIDs(allStudentID))
            {
                if (obj.ReceiveAddress == "戶籍地址")
                {
                    if (!string.IsNullOrEmpty(record.PermanentAddress))
                        StudentSuperOBJ[record.RefStudentID].address = record.Permanent.County + record.Permanent.Town + record.Permanent.District + record.Permanent.Area + record.Permanent.Detail;

                    if (!string.IsNullOrEmpty(record.PermanentZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.PermanentZipCode;

                        if (record.PermanentZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.PermanentZipCode.Substring(0, 1);
                        if (record.PermanentZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.PermanentZipCode.Substring(1, 1);
                        if (record.PermanentZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.PermanentZipCode.Substring(2, 1);
                        if (record.PermanentZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.PermanentZipCode.Substring(3, 1);
                        if (record.PermanentZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.PermanentZipCode.Substring(4, 1);
                    }

                }
                else if (obj.ReceiveAddress == "聯絡地址")
                {
                    if (!string.IsNullOrEmpty(record.MailingAddress))
                        StudentSuperOBJ[record.RefStudentID].address = record.Mailing.County + record.Mailing.Town + record.Mailing.District + record.Mailing.Area + record.Mailing.Detail; //再處理

                    if (!string.IsNullOrEmpty(record.MailingZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.MailingZipCode;

                        if (record.MailingZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.MailingZipCode.Substring(0, 1);
                        if (record.MailingZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.MailingZipCode.Substring(1, 1);
                        if (record.MailingZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.MailingZipCode.Substring(2, 1);
                        if (record.MailingZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.MailingZipCode.Substring(3, 1);
                        if (record.MailingZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.MailingZipCode.Substring(4, 1);
                    }
                }
                else if (obj.ReceiveAddress == "其他地址")
                {
                    if (!string.IsNullOrEmpty(record.Address1Address))
                        StudentSuperOBJ[record.RefStudentID].address = record.Address1.County + record.Address1.Town + record.Address1.District + record.Address1.Area + record.Address1.Detail; //再處理

                    if (!string.IsNullOrEmpty(record.Address1ZipCode))
                    {
                        StudentSuperOBJ[record.RefStudentID].ZipCode = record.Address1ZipCode;

                        if (record.Address1ZipCode.Length >= 1)
                            StudentSuperOBJ[record.RefStudentID].ZipCode1 = record.Address1ZipCode.Substring(0, 1);
                        if (record.Address1ZipCode.Length >= 2)
                            StudentSuperOBJ[record.RefStudentID].ZipCode2 = record.Address1ZipCode.Substring(1, 1);
                        if (record.Address1ZipCode.Length >= 3)
                            StudentSuperOBJ[record.RefStudentID].ZipCode3 = record.Address1ZipCode.Substring(2, 1);
                        if (record.Address1ZipCode.Length >= 4)
                            StudentSuperOBJ[record.RefStudentID].ZipCode4 = record.Address1ZipCode.Substring(3, 1);
                        if (record.Address1ZipCode.Length >= 5)
                            StudentSuperOBJ[record.RefStudentID].ZipCode5 = record.Address1ZipCode.Substring(4, 1);
                    }
                }
            }
            _BGWAbsenceNotification.ReportProgress(70);
            #endregion

            #region 取得學生監護人父母親資料
            foreach (ParentRecord record in Parent.SelectByStudentIDs(allStudentID))
            {
                StudentSuperOBJ[record.RefStudentID].CustodianName = record.CustodianName;
                StudentSuperOBJ[record.RefStudentID].FatherName = record.FatherName;
                StudentSuperOBJ[record.RefStudentID].MotherName = record.MotherName;
            }
            //dsrsp = JHSchool.Compatibility.Feature.QueryStudent.GetMultiParentInfo(allStudentID.ToArray());
            //foreach (XmlElement var in dsrsp.GetContent().GetElements("ParentInfo"))
            //{
            //    string studentID = var.GetAttribute("StudentID");

            //    studentInfo[studentID].Add("CustodianName", var.SelectSingleNode("CustodianName").InnerText);
            //    studentInfo[studentID].Add("FatherName", var.SelectSingleNode("FatherName").InnerText);
            //    studentInfo[studentID].Add("MotherName", var.SelectSingleNode("MotherName").InnerText);
            //}

            _BGWAbsenceNotification.ReportProgress(80);
            #endregion

            #endregion

            #region 通用資料

            Allmapping.Add("學年度", obj.SchoolYear);
            Allmapping.Add("學期", obj.Semester);
            Allmapping.Add("試別", obj.ExamName);
            Allmapping.Add("缺曠獎懲統計期間", obj.StartDate.ToShortDateString() + " 至 " + obj.EndDate.ToShortDateString());
            Allmapping.Add("校長", K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText);
            Allmapping.Add("教務主任", K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("EduDirectorName").InnerText);


            Allmapping.Add("科目名稱1", "");
            Allmapping.Add("科目名稱2", "");
            Allmapping.Add("科目名稱3", "");
            Allmapping.Add("科目名稱4", "");
            Allmapping.Add("科目名稱5", "");
            Allmapping.Add("科目節數1", "");
            Allmapping.Add("科目節數2", "");
            Allmapping.Add("科目節數3", "");
            Allmapping.Add("科目節數4", "");
            Allmapping.Add("科目節數5", "");
            Allmapping.Add("CN", "");
            Allmapping.Add("POSTALCODE", "");
            Allmapping.Add("POSTALADDRESS", "");
            Allmapping.Add("學號", "");
            Allmapping.Add("班級", "");
            Allmapping.Add("座號", "");
            Allmapping.Add("學生姓名", "");
            Allmapping.Add("成績1", "");
            Allmapping.Add("成績2", "");
            Allmapping.Add("成績3", "");
            Allmapping.Add("成績4", "");
            Allmapping.Add("成績5", "");
            Allmapping.Add("加權平均", "");
            Allmapping.Add("加權總分", "");
            Allmapping.Add("名次", "");
            Allmapping.Add("年排名", "");
            Allmapping.Add("平時成績1", "");
            Allmapping.Add("平時成績2", "");
            Allmapping.Add("平時成績3", "");
            Allmapping.Add("平時成績4", "");
            Allmapping.Add("平時成績5", "");
            Allmapping.Add("平時加權平均", "");
            Allmapping.Add("評量總成績1", "");
            Allmapping.Add("評量總成績2", "");
            Allmapping.Add("評量總成績3", "");
            Allmapping.Add("評量總成績4", "");
            Allmapping.Add("評量總成績5", "");
            Allmapping.Add("評量總加權平均", "");
            Allmapping.Add("大功", "0");
            Allmapping.Add("小功", "0");
            Allmapping.Add("嘉獎", "0");
            Allmapping.Add("大過", "0");
            Allmapping.Add("小過", "0");
            Allmapping.Add("警告", "0");
            Allmapping.Add("曠課", "0");
            Allmapping.Add("事假", "0");
            Allmapping.Add("病假", "0");
            Allmapping.Add("喪假", "0");
            Allmapping.Add("公假", "0");
            Allmapping.Add("家長代碼", "");

            #endregion

            //2019/12/05 穎驊註解，弘文epost 這個case 比較特別，經過業務PM嘉詮確認，它們學校是已經跟郵局談好CSV檔的格式(先前都是行政手動產生)
            //然後我們要來接，它一份CSV檔的格式 同時包括了定期、學期資料，因此將本程式設計模式，能夠依不同的需求分別產生對應的資料
            // EX: 產生定期資料時，學期資料資料就是空白
            #region 學期資料
            Allmapping.Add("國文百分成績", "");
            Allmapping.Add("英語百分成績", "");
            Allmapping.Add("數學百分成績", "");
            Allmapping.Add("社會百分成績", "");
            Allmapping.Add("自然科學百分成績", "");
            Allmapping.Add("理化百分成績", "");
            Allmapping.Add("自然百分成績", "");
            Allmapping.Add("資訊科技百分成績", "");
            Allmapping.Add("生活科技百分成績", "");
            Allmapping.Add("音樂百分成績", "");
            Allmapping.Add("視覺藝術百分成績", "");
            Allmapping.Add("表演藝術百分成績", "");
            Allmapping.Add("家政百分成績", "");
            Allmapping.Add("童軍百分成績", "");
            Allmapping.Add("輔導百分成績", "");
            Allmapping.Add("健康教育百分成績", "");
            Allmapping.Add("體育百分成績", "");
            Allmapping.Add("英語聽講百分成績", "");
            Allmapping.Add("資訊應用百分成績", "");
            Allmapping.Add("ESL百分成績", "");
            Allmapping.Add("地球科學百分成績", "");
            Allmapping.Add("閱讀理解百分成績", "");
            Allmapping.Add("閱讀與寫作百分成績", "");
            Allmapping.Add("語文表達百分成績", "");
            Allmapping.Add("國文節數", "");
            Allmapping.Add("英語節數", "");
            Allmapping.Add("數學節數", "");
            Allmapping.Add("社會節數", "");
            Allmapping.Add("自然科學節數", "");
            Allmapping.Add("理化節數", "");
            Allmapping.Add("自然節數", "");
            Allmapping.Add("資訊科技節數", "");
            Allmapping.Add("生活科技節數", "");
            Allmapping.Add("音樂節數", "");
            Allmapping.Add("視覺藝術節數", "");
            Allmapping.Add("表演藝術節數", "");
            Allmapping.Add("家政節數", "");
            Allmapping.Add("童軍節數", "");
            Allmapping.Add("輔導節數", "");
            Allmapping.Add("健康教育節數", "");
            Allmapping.Add("體育節數", "");
            Allmapping.Add("英語聽講節數", "");
            Allmapping.Add("資訊應用節數", "");
            Allmapping.Add("ESL節數", "");
            Allmapping.Add("地球科學節數", "");
            Allmapping.Add("閱讀理解節數", "");
            Allmapping.Add("閱讀與寫作節數", "");
            Allmapping.Add("語文表達節數", "");
            Allmapping.Add("國文等第", "");
            Allmapping.Add("英語等第", "");
            Allmapping.Add("數學等第", "");
            Allmapping.Add("社會等第", "");
            Allmapping.Add("自然科學等第", "");
            Allmapping.Add("理化等第", "");
            Allmapping.Add("自然等第", "");
            Allmapping.Add("資訊科技等第", "");
            Allmapping.Add("生活科技等第", "");
            Allmapping.Add("音樂等第", "");
            Allmapping.Add("視覺藝術等第", "");
            Allmapping.Add("表演藝術等第", "");
            Allmapping.Add("家政等第", "");
            Allmapping.Add("童軍等第", "");
            Allmapping.Add("輔導等第", "");
            Allmapping.Add("健康教育等第", "");
            Allmapping.Add("體育等第", "");
            Allmapping.Add("英語聽講等第", "");
            Allmapping.Add("資訊應用等第", "");
            Allmapping.Add("ESL等第", "");
            Allmapping.Add("地球科學等第", "");
            Allmapping.Add("閱讀理解等第", "");
            Allmapping.Add("閱讀與寫作等第", "");
            Allmapping.Add("語文表達等第", "");
            Allmapping.Add("國文文字描述", "");
            Allmapping.Add("英語文字描述", "");
            Allmapping.Add("數學文字描述", "");
            Allmapping.Add("社會文字描述", "");
            Allmapping.Add("自然科學文字描述", "");
            Allmapping.Add("理化文字描述", "");
            Allmapping.Add("自然文字描述", "");
            Allmapping.Add("資訊科技文字描述", "");
            Allmapping.Add("生活科技文字描述", "");
            Allmapping.Add("音樂文字描述", "");
            Allmapping.Add("視覺藝術文字描述", "");
            Allmapping.Add("表演藝術文字描述", "");
            Allmapping.Add("家政文字描述", "");
            Allmapping.Add("童軍文字描述", "");
            Allmapping.Add("輔導文字描述", "");
            Allmapping.Add("健康教育文字描述", "");
            Allmapping.Add("體育文字描述", "");
            Allmapping.Add("英語聽講文字描述", "");
            Allmapping.Add("資訊應用文字描述", "");
            Allmapping.Add("ESL文字描述", "");
            Allmapping.Add("地球科學文字描述", "");
            Allmapping.Add("閱讀理解文字描述", "");
            Allmapping.Add("閱讀與寫作文字描述", "");
            Allmapping.Add("語文表達文字描述", "");
            Allmapping.Add("學期大功", "");
            Allmapping.Add("學期小功", "");
            Allmapping.Add("學期嘉獎", "");
            Allmapping.Add("學期大過", "");
            Allmapping.Add("學期小過", "");
            Allmapping.Add("學期警告", "");
            Allmapping.Add("學期曠課", "");
            Allmapping.Add("學期事假", "");
            Allmapping.Add("學期病假", "");
            Allmapping.Add("學期喪假", "");
            Allmapping.Add("學期公假", "");
            Allmapping.Add("學期遲到", "");
            Allmapping.Add("愛整潔", "");
            Allmapping.Add("有禮貌", "");
            Allmapping.Add("守秩序", "");
            Allmapping.Add("責任心", "");
            Allmapping.Add("公德心", "");
            Allmapping.Add("友愛關懷", "");
            Allmapping.Add("團隊合作", "");
            Allmapping.Add("團體活動表現", "");
            Allmapping.Add("導師評語", "");
            Allmapping.Add("服務學習時數", "");
            Allmapping.Add("科目班級平均", "");
            Allmapping.Add("科目排名", "");
            Allmapping.Add("班級加權總分", "");
            Allmapping.Add("班級加權平均", "");
            Allmapping.Add("學期科目班級平均", "");
            Allmapping.Add("學期科目排名", "");
            Allmapping.Add("學期班級加權總分", "");
            Allmapping.Add("學期班級加權平均", "");
            Allmapping.Add("班級人數", "");
            Allmapping.Add("年級人數", "");
            Allmapping.Add("科目PR值", "");
            #endregion

            #region 取得學生成績計算規則
            ScoreCalculator defaultScoreCalculator = new ScoreCalculator(null);

            //key: ScoreCalcRuleID
            Dictionary<string, ScoreCalculator> calcCache = new Dictionary<string, ScoreCalculator>();
            //key: StudentID, val: ScoreCalcRuleID
            Dictionary<string, string> calcIDCache = new Dictionary<string, string>();
            List<string> scoreCalcRuleIDList = new List<string>();
            foreach (StudentRecord student in SelectedStudents)
            {
                //calcCache.Add(student.ID, new ScoreCalculator(student.ScoreCalcRule));
                string calcID = string.Empty;
                if (!string.IsNullOrEmpty(student.OverrideScoreCalcRuleID))
                    calcID = student.OverrideScoreCalcRuleID;
                else if (student.Class != null && !string.IsNullOrEmpty(student.Class.RefScoreCalcRuleID))
                    calcID = student.Class.RefScoreCalcRuleID;

                if (!string.IsNullOrEmpty(calcID))
                    calcIDCache.Add(student.ID, calcID);
            }
            foreach (JHScoreCalcRuleRecord record in JHScoreCalcRule.SelectByIDs(calcIDCache.Values))
            {
                if (!calcCache.ContainsKey(record.ID))
                    calcCache.Add(record.ID, new ScoreCalculator(record));
            }

            #endregion

            #region 取得定期評量資料
            // 課程資料
            Dictionary<string, JHCourseRecord> CourseDict = new Dictionary<string, JHCourseRecord>();

            foreach (JHCourseRecord co in JHCourse.SelectBySchoolYearAndSemester(int.Parse(obj.SchoolYear), int.Parse(obj.Semester)))
            {
                CourseDict.Add(co.ID, co);
            }

            // 取評量成績
            Dictionary<string, List<HC.JHSCETakeRecord>> Score1Dict = new Dictionary<string, List<HC.JHSCETakeRecord>>();
            foreach (JHSCETakeRecord record in JHSCETake.SelectByStudentAndCourse(allStudentID, CourseDict.Keys.ToList()))
            {
                if (record.RefExamID == obj.ExamID)
                {
                    if (!Score1Dict.ContainsKey(record.RefStudentID))
                        Score1Dict.Add(record.RefStudentID, new List<HC.JHSCETakeRecord>());

                    Score1Dict[record.RefStudentID].Add(new HC.JHSCETakeRecord(record));
                }
            }

            // 取得這次該修課程
            Dictionary<string, Dictionary<string, DAO.SubjectDomainName>> StudCourseDict = Utility.GetStudentSCAttendCourse(allStudentID, CourseDict.Keys.ToList(), obj.ExamID);

            // 取得評量設定比例
            Dictionary<string, decimal> ScorePercentageHSDict = Utility.GetScorePercentageHS();

            // 處理評量成績科目
            Dictionary<string, DAO.StudExamScore> studExamScoreDict = new Dictionary<string, DAO.StudExamScore>();
            foreach (string studID in allStudentID)
            {
                // 成績計算規則
                ScoreCalculator studentCalculator = defaultScoreCalculator;
                if (calcIDCache.ContainsKey(studID) && calcCache.ContainsKey(calcIDCache[studID]))
                    studentCalculator = calcCache[calcIDCache[studID]];

                if (Score1Dict.ContainsKey(studID))
                {
                    if (!studExamScoreDict.ContainsKey(studID))
                        studExamScoreDict.Add(studID, new DAO.StudExamScore(studentCalculator));

                    foreach (HC.JHSCETakeRecord rec in Score1Dict[studID])
                    {
                        if (rec.RefExamID == obj.ExamID && CourseDict.ContainsKey(rec.RefCourseID))
                        {
                            JHCourseRecord cr = CourseDict[rec.RefCourseID];

                            string SubjecName = cr.Subject;

                            // 勾選科目
                            if (obj.SelSubjNameList.Contains(SubjecName))
                            {
                                if (!studExamScoreDict[studID]._ExamSubjectScoreDict.ContainsKey(SubjecName))
                                {
                                    DAO.ExamSubjectScore ess = new DAO.ExamSubjectScore();
                                    ess.DomainName = cr.Domain;
                                    ess.SubjectName = SubjecName;
                                    ess.ScoreA = rec.AssignmentScore;
                                    ess.ScoreF = rec.Score;

                                    if (ess.ScoreA.HasValue && ess.ScoreF.HasValue)
                                    {
                                        if (ScorePercentageHSDict.ContainsKey(cr.RefAssessmentSetupID))
                                        {
                                            // 取得定期，評量由100-定期
                                            decimal f = ScorePercentageHSDict[cr.RefAssessmentSetupID] * 0.01M;
                                            decimal a = (100 - ScorePercentageHSDict[cr.RefAssessmentSetupID]) * 0.01M;
                                            ess.ScoreT = ess.ScoreA.Value * a + ess.ScoreF.Value * f;
                                        }
                                        else
                                            ess.ScoreT = ess.ScoreA.Value * 0.5M + ess.ScoreF.Value * 0.5M; // 沒有設定預設50,50

                                        // 原本
                                        //ess.ScoreT = (ess.ScoreA.Value + ess.ScoreF.Value) / 2;
                                    }
                                    if (ess.ScoreA.HasValue && ess.ScoreF.HasValue == false)
                                        ess.ScoreT = ess.ScoreA.Value;

                                    if (ess.ScoreA.HasValue == false && ess.ScoreF.HasValue)
                                        ess.ScoreT = ess.ScoreF.Value;

                                    // 依照成績計算規則科目方式處理進位，只有總成績。
                                    // 平時
                                    //if(ess.ScoreA.HasValue)
                                    //    ess.ScoreA = studentCalculator.ParseSubjectScore(ess.ScoreA.Value);

                                    // 定期
                                    //if (ess.ScoreF.HasValue)
                                    //    ess.ScoreF = studentCalculator.ParseSubjectScore(ess.ScoreF.Value);

                                    if (ess.ScoreT.HasValue)
                                        ess.ScoreT = studentCalculator.ParseSubjectScore(ess.ScoreT.Value);

                                    ess.Text = rec.Text;
                                    ess.Credit = cr.Credit;
                                    studExamScoreDict[studID]._ExamSubjectScoreDict.Add(SubjecName, ess);
                                }
                            }
                        }
                    }
                    // 計算領域成績
                    studExamScoreDict[studID].CalcSubjectToDomain();
                }

                if (StudCourseDict.ContainsKey(studID))
                {
                    if (!studExamScoreDict.ContainsKey(studID))
                        studExamScoreDict.Add(studID, new DAO.StudExamScore(studentCalculator));
                    if (studExamScoreDict[studID]._ExamSubjectScoreDict == null)
                    {
                        studExamScoreDict[studID]._ExamSubjectScoreDict = new Dictionary<string, DAO.ExamSubjectScore>();
                        studExamScoreDict[studID]._ExamDomainScoreDict = new Dictionary<string, DAO.ExamDomainScore>();
                    }
                    foreach (KeyValuePair<string, DAO.SubjectDomainName> data in StudCourseDict[studID])
                    {
                        // 沒有勾選不加入
                        if (!obj.SelSubjNameList.Contains(data.Key))
                            continue;

                        // 加入有修課沒有成績空科目
                        if (!studExamScoreDict[studID]._ExamSubjectScoreDict.ContainsKey(data.Key))
                        {
                            DAO.ExamSubjectScore ess = new DAO.ExamSubjectScore();
                            ess.SubjectName = data.Key;
                            ess.DomainName = data.Value.DomainName;
                            ess.Credit = data.Value.Credit;
                            studExamScoreDict[studID]._ExamSubjectScoreDict.Add(data.Key, ess);
                        }
                    }
                }

            }

            _BGWAbsenceNotification.ReportProgress(90);
            #endregion

            #region 取得固定排名資料

            //抓取固定排名資料
            string sql_rank = @"
SELECT 
	rank_matrix.id AS rank_matrix_id
	, rank_matrix.school_year
	, rank_matrix.semester
	, rank_matrix.grade_year
	, rank_matrix.item_type
	, rank_matrix.ref_exam_id
	, rank_matrix.item_name
	, rank_matrix.rank_type
	, rank_matrix.rank_name
	, class.class_name
	, student.seat_no
	, student.student_number
	, student.name
	, rank_detail.ref_student_id
    , rank_detail.score
	, rank_detail.rank
	, rank_detail.pr
	, rank_detail.percentile
    , rank_matrix.avg_top_25
    , rank_matrix.avg_top_50
    , rank_matrix.avg
    , rank_matrix.avg_bottom_50
    , rank_matrix.avg_bottom_25
FROM 
	rank_matrix
	LEFT OUTER JOIN rank_detail
		ON rank_detail.ref_matrix_id = rank_matrix.id
	LEFT OUTER JOIN student
		ON student.id = rank_detail.ref_student_id
	LEFT OUTER JOIN class
		ON class.id = student.ref_class_id
WHERE
	rank_matrix.is_alive = true
	AND rank_matrix.school_year = '" + obj.SchoolYear + @"'
    AND rank_matrix.semester = '" + obj.Semester + @"'
	AND rank_matrix.item_type like '定期評量%'
	AND rank_matrix.ref_exam_id = '" + obj.ExamID + @"'
    AND ref_student_id IN ('" + string.Join("','", allStudentID) + @"') 
ORDER BY 
	rank_matrix.id
	, rank_detail.rank
	, class.grade_year
	, class.display_order
	, class.class_name
	, student.seat_no
	, student.id";

            QueryHelper qh = new QueryHelper();

            DataTable datatableRank = qh.Select(sql_rank);

            // 處理 固定排名 (2019/12/10目前沒有平時評量的固定排名，與業務嘉詮佳樺討論後，他們說先暫時不用 )
            Dictionary<string, Dictionary<string, string>> studScoreRankDict = new Dictionary<string, Dictionary<string, string>>();

            foreach (DataRow dr in datatableRank.Rows)
            {
                if (!studScoreRankDict.ContainsKey("" + dr["ref_student_id"]))
                {
                    studScoreRankDict.Add("" + dr["ref_student_id"], new Dictionary<string, string>());

                    if ("" + dr["item_type"] == "定期評量/總計成績" && "" + dr["item_name"] == "加權平均")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("評量總加權平均"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("評量總加權平均", "" + dr["score"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權平均")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("加權平均"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("加權平均", "" + dr["score"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權總分")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("加權總分"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("加權總分", "" + dr["score"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權總分" && "" + dr["rank_type"] == "班排名")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("名次"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("名次", "" + dr["rank"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權總分" && "" + dr["rank_type"] == "年排名")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("年排名"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("年排名", "" + dr["rank"]);
                        }
                    }

                }
                else
                {
                    if ("" + dr["item_type"] == "定期評量/總計成績" && "" + dr["item_name"] == "加權平均")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("評量總加權平均"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("評量總加權平均", "" + dr["score"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權平均")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("加權平均"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("加權平均", "" + dr["score"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權總分")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("加權總分"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("加權總分", "" + dr["score"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權總分" && "" + dr["rank_type"] == "班排名")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("名次"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("名次", "" + dr["rank"]);
                        }
                    }

                    if ("" + dr["item_type"] == "定期評量_定期/總計成績" && "" + dr["item_name"] == "加權總分" && "" + dr["rank_type"] == "年排名")
                    {
                        if (!studScoreRankDict["" + dr["ref_student_id"]].ContainsKey("年排名"))
                        {
                            studScoreRankDict["" + dr["ref_student_id"]].Add("年排名", "" + dr["rank"]);
                        }
                    }
                }
            }

            _BGWAbsenceNotification.ReportProgress(100);
            #endregion


            #region 產生報表

            DataTable dt = new DataTable();

            foreach (string studentID in StudentSuperOBJ.Keys)
            {

                StudentOBJ eachStudentInfo = StudentSuperOBJ[studentID];

                //合併列印的資料
                Dictionary<string, object> mapping = new Dictionary<string, object>();

                if (!DelStudent.Contains(studentID)) //如果不包含在內,就離開
                    continue;

                if (obj.PrintHasRecordOnly)
                {
                    //明細等於0
                    if (eachStudentInfo.studentAbsenceDetail.Count == 0)
                    {
                        currentStudentCount++;
                        continue;
                    }
                }

                // 弘文高中國中部 CSV 規格 沒有要用到這些
                ////學校資訊
                //mapping.Add("學校名稱", School.ChineseName);
                //mapping.Add("學校地址", School.Address);
                //mapping.Add("學校電話", School.Telephone);

                //學生資料
                mapping.Add("學生姓名", eachStudentInfo.student.Name);
                mapping.Add("班級", eachStudentInfo.ClassName);
                mapping.Add("座號", eachStudentInfo.SeatNo);
                mapping.Add("學號", eachStudentInfo.StudentNumber);
                mapping.Add("班級導師", eachStudentInfo.TeacherName);
                mapping.Add("資料期間", obj.StartDate.ToShortDateString() + " 至 " + obj.EndDate.ToShortDateString());

                // 2019/11/12 穎驊註解 本專案為弘文於本學期提出來的需求，增加家長代碼
                mapping.Add("家長代碼", eachStudentInfo.ParentCode);

                //收件人資料
                //if (obj.ReceiveName == "監護人姓名")
                //    mapping.Add("收件人姓名", eachStudentInfo.CustodianName);
                //else if (obj.ReceiveName == "父親姓名")
                //    mapping.Add("收件人姓名", eachStudentInfo.FatherName);
                //else if (obj.ReceiveName == "母親姓名")
                //    mapping.Add("收件人姓名", eachStudentInfo.MotherName);
                //else
                //    mapping.Add("收件人姓名", eachStudentInfo.student.Name);

                if (obj.ReceiveName == "監護人姓名")
                    mapping.Add("CN", eachStudentInfo.CustodianName);
                else if (obj.ReceiveName == "父親姓名")
                    mapping.Add("CN", eachStudentInfo.FatherName);
                else if (obj.ReceiveName == "母親姓名")
                    mapping.Add("CN", eachStudentInfo.MotherName);
                else
                    mapping.Add("CN", eachStudentInfo.student.Name);

                //收件人地址資料
                //mapping.Add("收件人地址", eachStudentInfo.address);
                //mapping.Add("郵遞區號", eachStudentInfo.ZipCode);

                mapping.Add("POSTALADDRESS", eachStudentInfo.address);
                mapping.Add("POSTALCODE", eachStudentInfo.ZipCode);

                //mapping.Add("0", eachStudentInfo.ZipCode1);
                //mapping.Add("1", eachStudentInfo.ZipCode2);
                //mapping.Add("2", eachStudentInfo.ZipCode3);
                //mapping.Add("4", eachStudentInfo.ZipCode4);
                //mapping.Add("5", eachStudentInfo.ZipCode5);

                mapping.Add("學年度", School.DefaultSchoolYear);
                mapping.Add("學期", School.DefaultSemester);

                // 作為統計全部缺曠
                Dictionary<string, int> absenceTotalDict = new Dictionary<string, int>();

                //缺曠學期統計部份                          
                foreach (string attendanceType in configkeylist)
                {

                    foreach (string absenceType in config[attendanceType])
                    {
                        int dataValue = 0;
                        int semesterDataValue = 0;
                        string PeriodAndAbsence = attendanceType + "," + absenceType;
                        //本期統計
                        if (eachStudentInfo.studentAbsence.ContainsKey(PeriodAndAbsence))
                        {
                            dataValue = eachStudentInfo.studentAbsence[PeriodAndAbsence];
                        }
                        //學期統計
                        if (eachStudentInfo.studentSemesterAbsence.ContainsKey(PeriodAndAbsence))
                        {
                            semesterDataValue = eachStudentInfo.studentSemesterAbsence[PeriodAndAbsence];
                        }

                        if (!absenceTotalDict.ContainsKey(absenceType))
                        {
                            absenceTotalDict.Add(absenceType, dataValue);
                        }
                        else
                        {
                            absenceTotalDict[absenceType] += dataValue;
                        }

                        if (!absenceTotalDict.ContainsKey("學期" + absenceType))
                        {
                            absenceTotalDict.Add("學期" + absenceType, semesterDataValue);
                        }
                        else
                        {
                            absenceTotalDict["學期" + absenceType] += semesterDataValue;
                        }

                    }
                }

                foreach (string absence in absenceTotalDict.Keys)
                {
                    mapping.Add(absence, "" + absenceTotalDict[absence]);
                }

                // 獎懲統計
                foreach (string merit in meritNameList)
                {
                    if (eachStudentInfo.studentMerit.ContainsKey(merit))
                    {
                        mapping.Add(merit, eachStudentInfo.studentMerit[merit]);
                    }

                    if (eachStudentInfo.studentSemesterMerit.ContainsKey(merit))
                    {
                        mapping.Add("學期" + merit, eachStudentInfo.studentSemesterMerit[merit]);
                    }

                }

                // 定期評量成績
                if (studExamScoreDict.ContainsKey(studentID))
                {
                    int subjectIndex = 1;

                    foreach (string subject in obj.SelSubjNameList)
                    {
                        // 2019/12/10 弘文國中開的科目數，以五個為上限
                        if (subjectIndex >= 5)
                        {
                            break;
                        }
                        if (studExamScoreDict[studentID]._ExamSubjectScoreDict.ContainsKey(subject))
                        {
                            mapping.Add("科目名稱" + subjectIndex, studExamScoreDict[studentID]._ExamSubjectScoreDict[subject].SubjectName);

                            mapping.Add("科目節數" + subjectIndex, studExamScoreDict[studentID]._ExamSubjectScoreDict[subject].Credit);

                            mapping.Add("成績" + subjectIndex, studExamScoreDict[studentID]._ExamSubjectScoreDict[subject].ScoreF);

                            mapping.Add("平時成績" + subjectIndex, studExamScoreDict[studentID]._ExamSubjectScoreDict[subject].ScoreA);

                            mapping.Add("評量總成績" + subjectIndex, studExamScoreDict[studentID]._ExamSubjectScoreDict[subject].ScoreT);
                        }
                        subjectIndex++;
                    }

                }

                //固定排名
                if (studScoreRankDict.ContainsKey(studentID))
                {
                    foreach (string rankName in studScoreRankDict[studentID].Keys)
                    {
                        mapping.Add(rankName, studScoreRankDict[studentID][rankName]);
                    }
                }


                #region epost 使用

                // 將對應功能變數 套入dt
                foreach (string key in Allmapping.Keys)
                {
                    if (!dt.Columns.Contains(key))
                    {
                        dt.Columns.Add(key);
                    }
                }

                foreach (string key in mapping.Keys)
                {
                    if (!dt.Columns.Contains(key))
                    {
                        dt.Columns.Add(key);
                    }
                }

                DataRow row = dt.NewRow();

                foreach (string key in Allmapping.Keys)
                {
                    row[key] = Allmapping[key];
                }

                foreach (string key in mapping.Keys)
                {
                    row[key] = mapping[key];
                }

                dt.Rows.Add(row);
                #endregion

                //回報進度
                _BGWAbsenceNotification.ReportProgress((int)(((double)currentStudentCount++ * 100.0) / (double)totalStudentNumber));
            }

            #endregion


            e.Result = new object[] { dt };
        }

        /// <summary>
        /// 移除
        /// </summary>
        private string Switching(string abc)
        {
            if (!string.IsNullOrEmpty(abc))
            {
                string[] splitDate = abc.Split('/');
                return splitDate[1] + "/" + splitDate[2];
            }
            else
                return "";
        }
    }
}

