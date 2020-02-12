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

namespace K12.缺曠通知單2015
{
    internal class Report : IReport
    {
        private BackgroundWorker _BGWAbsenceNotification;

        private List<StudentRecord> SelectedStudents { get; set; }

        private Dictionary<string, List<string>> config;

        private static QueryHelper queryHelper;

        string addconfigName = "缺曠通知單_ForK12_缺曠別設定.2013_弘文epost";

        List<string> configkeylist { get; set; }
        ConfigOBJ obj;

        string entityName;

        /// <summary>
        /// 學生電子報表
        /// </summary>
        SmartSchool.ePaper.ElectronicPaper paperForStudent { get; set; }

        //轉縮寫
        Dictionary<string, string> absenceList = new Dictionary<string, string>();

        public Report(string _entityName)
        {
            entityName = _entityName;
        }

        public void Print()
        {
            #region IReport 成員

            AbsenceNotificationSelectDateRangeForm form = new AbsenceNotificationSelectDateRangeForm();

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

                FISCA.Presentation.MotherForm.SetStatusBarMessage("正在初始化缺曠通知單...");

                #region 建立設定檔

                obj = new ConfigOBJ();
                obj.StartDate = form.StartDate;
                obj.EndDate = form.EndDate;
                //2019/12/19 經由業務們測試討論後 決定 預設 為只產生有缺曠的名單，且不給更動
                //obj.PrintHasRecordOnly = form.PrintHasRecordOnly;

                obj.PrintHasRecordOnly = true;

                obj.Template = form.Template;
                //obj.userDefinedConfig = config;
                obj.ReceiveName = form.ReceiveName;
                obj.ReceiveAddress = form.ReceiveAddress;
                obj.ConditionName = form.ConditionName;
                obj.ConditionNumber = form.ConditionNumber;
                obj.ConditionName2 = form.ConditionName2;
                obj.ConditionNumber2 = form.ConditionNumber2;
                obj.PrintStudentList = form.PrintStudentList;
                obj.PaperUpdate = form._cbPaper; //是否列印電子報表

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

            #endregion
            string reportName = "缺曠通知單";

            #region 快取資料

            //超級資訊物件
            Dictionary<string, StudentOBJ> StudentSuperOBJ = new Dictionary<string, StudentOBJ>();

            //合併列印的資料
            Dictionary<string, object> Allmapping = new Dictionary<string, object>();
            Dictionary<string, string> ReversionDic = new Dictionary<string, string>();

            //所有學生ID
            List<string> allStudentID = new List<string>();

            //學生人數
            int currentStudentCount = 1;
            int totalStudentNumber = 0;

            #region 取得 Period List
            List<string> periodList = new List<string>();
            Dictionary<string, string> TestPeriodList = new Dictionary<string, string>();
            int PeriodX = 1;
            Dictionary<string, object> mappingAccessory_copy = new Dictionary<string, object>();
            foreach (K12.Data.PeriodMappingInfo each in K12.Data.PeriodMapping.SelectAll())
            {
                if (!periodList.Contains(each.Name))
                    periodList.Add(each.Name);

                if (!TestPeriodList.ContainsKey(each.Name)) //節次<-->類別
                    TestPeriodList.Add(each.Name, each.Type);

                Allmapping.Add("節次" + PeriodX, each.Name);
                mappingAccessory_copy.Add("節次" + PeriodX, each.Name);
                ReversionDic.Add(each.Name, "節次" + PeriodX);
                PeriodX++;
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

            //????使用者所選取的所有假別種類????
            List<string> userDefinedAbsenceList = new List<string>();

            int DefinedType = 1;
            foreach (string kind in configkeylist)
            {
                int DefinedAbsence = 1;
                Allmapping.Add("類型" + DefinedType, kind);

                foreach (string type in config[kind])
                {
                    Allmapping.Add("類型" + DefinedType + "缺曠" + DefinedAbsence, type);
                    Allmapping.Add("類型" + DefinedType + "縮寫" + DefinedAbsence, absenceList[type]);
                    DefinedAbsence++;

                    if (!userDefinedAbsenceList.Contains(type))
                    {
                        userDefinedAbsenceList.Add(type);
                    }
                }

                DefinedType++;
            }

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
            #endregion


            #region 取得所有學生缺曠紀錄，日期區間

            List<AttendanceRecord> attendanceList = K12.Data.Attendance.SelectByDate(SelectedStudents, obj.StartDate, obj.EndDate);

            if (attendanceList.Count == 0)
                e.Cancel = true; //沒有缺曠資料

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

            #endregion

            List<string> DelStudent = new List<string>(); //列印的學生

            #region 條件1
            if (obj.ConditionName != "") //如果不等於空就是要判斷啦
            {
                foreach (string each1 in StudentSuperOBJ.Keys) //取出一個學生
                {
                    int AbsenceCount = 0;
                    bool AbsenceBOOL = false;
                    foreach (string each2 in StudentSuperOBJ[each1].studentAbsenceDetail.Keys) //取出一天
                    {
                        foreach (string each3 in StudentSuperOBJ[each1].studentAbsenceDetail[each2].Keys) //取出一節內容
                        {
                            string each4 = StudentSuperOBJ[each1].studentAbsenceDetail[each2][each3];

                            if (TestPeriodList.ContainsKey(each3))
                            {
                                if (config.ContainsKey(TestPeriodList[each3]))
                                {
                                    if (obj.ConditionName == each4)
                                    {
                                        AbsenceCount++;
                                    }

                                    if (AbsenceCount >= int.Parse(obj.ConditionNumber))
                                    {
                                        AbsenceBOOL = true;
                                        if (!DelStudent.Contains(each1))
                                        {
                                            DelStudent.Add(each1); //把學生ID記下
                                        }
                                    }

                                    if (AbsenceBOOL)
                                        break;
                                }
                            }
                        }
                        if (AbsenceBOOL)
                            break;
                    }
                }
            }
            #endregion

            #region 條件2
            if (obj.ConditionName2 != "") //如果等於空就是直接全部印啦!!
            {
                foreach (string each1 in StudentSuperOBJ.Keys) //取出一個學生
                {
                    int AbsenceCount = 0;
                    bool AbsenceBOOL = false;
                    foreach (string each2 in StudentSuperOBJ[each1].studentAbsenceDetail.Keys) //取出一天
                    {
                        foreach (string each3 in StudentSuperOBJ[each1].studentAbsenceDetail[each2].Keys) //取出一節內容
                        {
                            string each4 = StudentSuperOBJ[each1].studentAbsenceDetail[each2][each3];

                            if (TestPeriodList.ContainsKey(each3))
                            {
                                if (config.ContainsKey(TestPeriodList[each3]))
                                {
                                    if (obj.ConditionName2 == each4)
                                    {
                                        AbsenceCount++;
                                    }

                                    if (AbsenceCount >= int.Parse(obj.ConditionNumber2))
                                    {
                                        AbsenceBOOL = true;

                                        DelStudent.Add(each1); //把學生ID記下
                                    }

                                    if (AbsenceBOOL)
                                        break;
                                }
                                if (AbsenceBOOL)
                                    break;
                            }
                        }
                    }
                }
            }
            #endregion

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
            foreach (AttendanceRecord attendance in K12.Data.Attendance.SelectBySchoolYearAndSemester(Student.SelectByIDs(allStudentID), int.Parse(School.DefaultSchoolYear), int.Parse(School.DefaultSemester)))
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
            #endregion

            #endregion

            Document template = new Document(obj.Template);

            #region 缺曠類別部份
            int columnNumber = 0;

            foreach (List<string> var in config.Values)
            {
                columnNumber += var.Count;
            }
            #endregion

            #region 產生報表

            Document doc = new Document();
            doc.Sections.Clear();
            paperForStudent = new SmartSchool.ePaper.ElectronicPaper("缺曠通知單_" + DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0'), School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);

            DataTable dt = new DataTable();
            
            // 2020/01/09 穎驊修正，因應郵局格式要求說明文件中，收件人的姓名、郵遞區號和地址，有規定的欄位名稱(CN,POSTALCODE,POSTALADDRESS)，手動填入
            dt.Columns.Add("CN");
            dt.Columns.Add("POSTALCODE");
            dt.Columns.Add("POSTALADDRESS");

            foreach (string studentID in StudentSuperOBJ.Keys)
            {
                Dictionary<string, object> mappingAccessory = new Dictionary<string, object>();
                foreach (string each in mappingAccessory_copy.Keys)
                {
                    mappingAccessory.Add(each, mappingAccessory_copy[each]);
                }

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

                Document eachSection = new Document();
                eachSection.Sections.Clear();
                eachSection.Sections.Add(eachSection.ImportNode(template.Sections[0], true));

                MemoryStream accessoryMemory;
                Aspose.Words.Document accessoryDoc;

                //學生資料
                mappingAccessory.Add("學生姓名", eachStudentInfo.student.Name);
                mappingAccessory.Add("班級", eachStudentInfo.ClassName);
                mappingAccessory.Add("座號", eachStudentInfo.SeatNo);
                mappingAccessory.Add("學號", eachStudentInfo.StudentNumber);
                mappingAccessory.Add("導師", eachStudentInfo.TeacherName);
                mappingAccessory.Add("學年度", School.DefaultSchoolYear);
                mappingAccessory.Add("學期", School.DefaultSemester);
                mappingAccessory.Add("資料期間", obj.StartDate.ToShortDateString() + " 至 " + obj.EndDate.ToShortDateString());
                //懲戒明細
                bool IsAccessory = false;

                //學校資訊
                mapping.Add("學校名稱", School.ChineseName);
                mapping.Add("學校地址", School.Address);
                mapping.Add("學校電話", School.Telephone);

                //學生資料
                mapping.Add("學生姓名", eachStudentInfo.student.Name);
                mapping.Add("班級", eachStudentInfo.ClassName);
                mapping.Add("座號", eachStudentInfo.SeatNo);
                mapping.Add("學號", eachStudentInfo.StudentNumber);
                mapping.Add("導師", eachStudentInfo.TeacherName);
                mapping.Add("資料期間", obj.StartDate.ToShortDateString() + " 至 " + obj.EndDate.ToShortDateString());

                // 2019/11/12 穎驊註解 本專案為弘文於本學期提出來的需求，增加家長代碼
                mapping.Add("家長代碼", eachStudentInfo.ParentCode);

                //收件人資料
                if (obj.ReceiveName == "監護人姓名")
                    mapping.Add("收件人姓名", eachStudentInfo.CustodianName);
                else if (obj.ReceiveName == "父親姓名")
                    mapping.Add("收件人姓名", eachStudentInfo.FatherName);
                else if (obj.ReceiveName == "母親姓名")
                    mapping.Add("收件人姓名", eachStudentInfo.MotherName);
                else
                    mapping.Add("收件人姓名", eachStudentInfo.student.Name);

                //收件人地址資料
                mapping.Add("收件人地址", eachStudentInfo.address);
                mapping.Add("郵遞區號", eachStudentInfo.ZipCode);
                mapping.Add("0", eachStudentInfo.ZipCode1);
                mapping.Add("1", eachStudentInfo.ZipCode2);
                mapping.Add("2", eachStudentInfo.ZipCode3);
                mapping.Add("4", eachStudentInfo.ZipCode4);
                mapping.Add("5", eachStudentInfo.ZipCode5);

                mapping.Add("學年度", School.DefaultSchoolYear);
                mapping.Add("學期", School.DefaultSemester);

                //缺曠學期統計部份
                int columnIndex = 1;
                int DefinedTypeCount = 1;
                foreach (string attendanceType in configkeylist)
                {
                    int DefinedAbsenceCount = 1;

                    foreach (string absenceType in config[attendanceType])
                    {
                        string dataValue = "0";
                        string semesterDataValue = "0";
                        string PeriodAndAbsence = attendanceType + "," + absenceType;
                        //本期統計
                        if (eachStudentInfo.studentAbsence.ContainsKey(PeriodAndAbsence))
                        {
                            dataValue = eachStudentInfo.studentAbsence[PeriodAndAbsence].ToString();
                        }
                        //學期統計
                        if (eachStudentInfo.studentSemesterAbsence.ContainsKey(PeriodAndAbsence))
                        {
                            semesterDataValue = eachStudentInfo.studentSemesterAbsence[PeriodAndAbsence].ToString();
                        }

                        mapping.Add("類型" + DefinedTypeCount + "本期" + DefinedAbsenceCount, dataValue);
                        mapping.Add("類型" + DefinedTypeCount + "學期" + DefinedAbsenceCount, semesterDataValue);
                        DefinedAbsenceCount++;
                        columnIndex++;
                    }
                    DefinedTypeCount++;
                }

                // 2020/02/07 嘉詮與學校後 確立 格式固定要有 日期1~日期12 且每日節次1~12，故總變數為日期1節次1 ~ 日期12節次12
                for (int date =1; date <= 12; date++)
                {
                    mapping.Add("日期" + date, "");

                    for (int peroid = 1; peroid <= 12; peroid++)
                    {
                        mapping.Add("日期" + date + "節次" + peroid, "");
                    }                   
                }


                //缺曠明細
                int DateCount = 1;
                foreach (string each in eachStudentInfo.studentAbsenceDetail.Keys)
                {
                    int Period = 1;
                    //資料數大於10,透過附件列印
                    if (DateCount <= 12)
                    {
                        if (mapping.ContainsKey("日期" + DateCount))
                        {
                            mapping["日期" + DateCount] = Switching(each);
                        }
                        else
                        {
                            mapping.Add("日期" + DateCount, Switching(each));
                        }                        
                    }

                    else {
                        mappingAccessory.Add("日期" + DateCount, Switching(each));
                    }
                    

                    //取得節次清單,一一檢查是否有資料要填
                    foreach (string Date in eachStudentInfo.studentAbsenceDetail[each].Keys)
                    {
                        string detail = eachStudentInfo.studentAbsenceDetail[each][Date];

                        if (absenceList.ContainsKey(detail))
                        {
                            if (ReversionDic.ContainsKey(Date))
                            {
                                if (DateCount <= 12) //資料數大於10,透過附件列印
                                {
                                    if (mapping.ContainsKey("日期" + DateCount + ReversionDic[Date]))
                                    {
                                        mapping["日期" + DateCount + ReversionDic[Date]] = absenceList[detail];
                                    }
                                    else
                                    {
                                        mapping.Add("日期" + DateCount + ReversionDic[Date],absenceList[detail]);
                                    }
                                    
                                    Period++;
                                }
                                else
                                {
                                    mappingAccessory.Add("日期" + DateCount + ReversionDic[Date], absenceList[detail]);
                                    IsAccessory = true;
                                    Period++;
                                }
                            }
                        }
                    }

                    DateCount++;
                }



                //學生個人資料
                string[] keys = new string[mapping.Count];
                object[] values = new object[mapping.Count];
                int i = 0;
                foreach (string key in mapping.Keys)
                {
                    keys[i] = key;
                    values[i++] = mapping[key];
                }
                eachSection.MailMerge.Execute(keys, values);


                //整體資料
                string[] Allkeys = new string[Allmapping.Count];
                object[] Allvalues = new object[Allmapping.Count];
                int t = 0;
                foreach (string key in Allmapping.Keys)
                {
                    Allkeys[t] = key;
                    Allvalues[t++] = Allmapping[key];
                }

                eachSection.MailMerge.Execute(Allkeys, Allvalues);
                eachSection.MailMerge.DeleteFields();

                #region epost 使用

               

                // 將對應功能變數 套入dt
                foreach (string key in mapping.Keys)
                {
                    if (!dt.Columns.Contains(key))
                    {
                        dt.Columns.Add(key);
                    }
                }

                foreach (string key in Allmapping.Keys)
                {
                    if (!dt.Columns.Contains(key))
                    {
                        dt.Columns.Add(key);
                    }
                }

                DataRow row = dt.NewRow();

                row["CN"] = mapping["收件人姓名"];
                row["POSTALCODE"] = mapping["郵遞區號"];
                row["POSTALADDRESS"] = mapping["收件人地址"];

                foreach (string key in mapping.Keys)
                {
                    row[key] = mapping[key];
                }

                foreach (string key in Allmapping.Keys)
                {
                    row[key] = Allmapping[key];
                }

                

                dt.Rows.Add(row);
                #endregion


                if (IsAccessory)
                {
                    accessoryMemory = new MemoryStream(Properties.Resources.缺曠通知單_附件一);
                    accessoryDoc = new Aspose.Words.Document(accessoryMemory);

                    string[] keysAccessory = new string[mappingAccessory.Count];
                    object[] valuesAccessory = new object[mappingAccessory.Count];
                    int xx = 0;
                    foreach (string key in mappingAccessory.Keys)
                    {
                        keysAccessory[xx] = key;
                        valuesAccessory[xx++] = mappingAccessory[key];
                    }

                    accessoryDoc.MailMerge.CleanupOptions = Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveEmptyParagraphs;
                    accessoryDoc.MailMerge.Execute(keysAccessory, valuesAccessory);
                    accessoryDoc.MailMerge.DeleteFields(); //刪除未合併之內容 

                    Aspose.Words.Node eachSectionaccessory = accessoryDoc.Sections[0].Clone();
                    eachSection.Sections.Add(eachSection.ImportNode(eachSectionaccessory, true));

                    MemoryStream stream = new MemoryStream();
                    eachSection.Save(stream, SaveFormat.Doc);
                    paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, eachStudentInfo.student.ID));
                }
                else
                {
                    MemoryStream stream = new MemoryStream();
                    eachSection.Save(stream, SaveFormat.Doc);
                    paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, eachStudentInfo.student.ID));
                }

                foreach (Aspose.Words.Section each in eachSection.Sections)
                {
                    Aspose.Words.Node eachSectionNode = each.Clone();
                    doc.Sections.Add(doc.ImportNode(eachSectionNode, true));
                }

                //回報進度
                _BGWAbsenceNotification.ReportProgress((int)(((double)currentStudentCount++ * 100.0) / (double)totalStudentNumber));
            }

            #endregion

            #region 產生學生清單

            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
            if (obj.PrintStudentList)
            {
                int CountRow = 0;
                wb.Worksheets[0].Cells[CountRow, 0].PutValue("班級");
                wb.Worksheets[0].Cells[CountRow, 1].PutValue("座號");
                wb.Worksheets[0].Cells[CountRow, 2].PutValue("學號");
                wb.Worksheets[0].Cells[CountRow, 3].PutValue("學生姓名");
                wb.Worksheets[0].Cells[CountRow, 4].PutValue("收件人姓名");
                wb.Worksheets[0].Cells[CountRow, 5].PutValue("地址");
                wb.Worksheets[0].Cells[CountRow, 6].PutValue("家長代碼");
                CountRow++;
                foreach (string each in StudentSuperOBJ.Keys)
                {
                    if (!DelStudent.Contains(each)) //如果不包含在內,就離開
                        continue;

                    if (obj.PrintHasRecordOnly)
                    {
                        //明細等於0
                        if (StudentSuperOBJ[each].studentAbsenceDetail.Count == 0)
                        {
                            currentStudentCount++;
                            continue;
                        }
                    }

                    wb.Worksheets[0].Cells[CountRow, 0].PutValue(StudentSuperOBJ[each].ClassName);
                    wb.Worksheets[0].Cells[CountRow, 1].PutValue(StudentSuperOBJ[each].SeatNo);
                    wb.Worksheets[0].Cells[CountRow, 2].PutValue(StudentSuperOBJ[each].StudentNumber);
                    wb.Worksheets[0].Cells[CountRow, 3].PutValue(StudentSuperOBJ[each].student.Name);
                    //收件人資料
                    if (obj.ReceiveName == "監護人姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].CustodianName);
                    else if (obj.ReceiveName == "父親姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].FatherName);
                    else if (obj.ReceiveName == "母親姓名")
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].MotherName);
                    else
                        wb.Worksheets[0].Cells[CountRow, 4].PutValue(StudentSuperOBJ[each].student.Name);

                    wb.Worksheets[0].Cells[CountRow, 5].PutValue(StudentSuperOBJ[each].ZipCode + " " + StudentSuperOBJ[each].address);
                    wb.Worksheets[0].Cells[CountRow, 6].PutValue(StudentSuperOBJ[each].ParentCode);
                    CountRow++;
                }
                wb.Worksheets[0].AutoFitColumns();
            }
            #endregion

            //是否上傳電子報表
            if (obj.PaperUpdate)
            {
                SmartSchool.ePaper.DispatcherProvider.Dispatch(paperForStudent);
            }

            string path = Path.Combine(Application.StartupPath, "Reports");
            string path2 = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");
            path2 = Path.Combine(path2, reportName + "(學生清單).xlsx");
            e.Result = new object[] { reportName, path, doc, path2, obj.PrintStudentList, wb, dt };
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
