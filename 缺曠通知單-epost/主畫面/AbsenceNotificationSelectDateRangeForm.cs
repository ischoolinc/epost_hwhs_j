using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using K12.Data.Configuration;

namespace K12.���m�q����2015
{
    public partial class AbsenceNotificationSelectDateRangeForm : SelectDateRangeForm
    {
        private MemoryStream _template = null;
        private MemoryStream _defaultTemplate = new MemoryStream(Properties.Resources.���m�q����_��}������);
        private byte[] _buffer = null;

        private bool _preferenceLoaded = false;
        private string _receiveName;
        private string _receiveAddress;
        private string _conditionName = "";
        private string _conditionNumber = "0";

        public string ConditionName { get { return _conditionName; } }
        public string ConditionNumber { get { return _conditionNumber; } }

        private string _conditionName2 = "";
        private string _conditionNumber2 = "0";

        public string ConditionName2 { get { return _conditionName2; } }
        public string ConditionNumber2 { get { return _conditionNumber2; } }

        string configName = "���m�q����_ForK12.2013_����epost";
        string addconfigName = "���m�q����_ForK12_���m�O�]�w.2013_����epost";

        public bool _cbPaper = false;

        public string ReceiveName
        {
            get { return _receiveName; }
        }
        public string ReceiveAddress
        {
            get { return _receiveAddress; }
        }

        public MemoryStream Template
        {
            get
            {
                if (_useDefaultTemplate)
                    return _defaultTemplate;
                else
                    return _template;
            }
        }

        //2019/12/19 �g�ѷ~�ȭ̴��հQ�׫� �M�w �w�] ���ۭq�A�B�������
        private AbsenceNotificationConfigForm.DateRangeMode _mode = AbsenceNotificationConfigForm.DateRangeMode.Custom;

        private bool _useDefaultTemplate = true;

        private bool _printHasRecordOnly = true;
        public bool PrintHasRecordOnly
        {
            get { return _printHasRecordOnly; }
        }

        //�O�_�C�L�ǥͲM��
        private bool _PrintStudentList = false;
        public bool PrintStudentList
        {
            get { return _PrintStudentList; }
        }

        public AbsenceNotificationSelectDateRangeForm()
        {
            InitializeComponent();
            Text = "���m�q����(����epost)";
            LoadPreference();
            InitialDateRange();
        }

        private void LoadPreference()
        {
            #region Ū�� Preference

            //XmlElement config = CurrentUser.Instance.Preference["���m�q����"];
            ConfigData cd = K12.Data.School.Configuration[configName];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config != null)
            {
                _useDefaultTemplate = bool.Parse(config.GetAttribute("Default"));

                XmlElement customize = (XmlElement)config.SelectSingleNode("CustomizeTemplate");
                XmlElement print = (XmlElement)config.SelectSingleNode("PrintHasRecordOnly");
                XmlElement dateRangeMode = (XmlElement)config.SelectSingleNode("DateRangeMode");
                XmlElement receive = (XmlElement)config.SelectSingleNode("Receive");
                XmlElement conditions = (XmlElement)config.SelectSingleNode("Conditions");
                XmlElement conditions2 = (XmlElement)config.SelectSingleNode("Conditions2");
                XmlElement PrintStudentList = (XmlElement)config.SelectSingleNode("PrintStudentList");

                if (customize != null)
                {
                    string templateBase64 = customize.InnerText;
                    _buffer = Convert.FromBase64String(templateBase64);
                    _template = new MemoryStream(_buffer);
                }

                if (print != null)
                {
                    if (print.HasAttribute("Checked"))
                    {
                        _printHasRecordOnly = bool.Parse(print.GetAttribute("Checked"));
                    }
                }
                else
                {
                    XmlElement newPrintHasRecordOnly = config.OwnerDocument.CreateElement("PrintHasRecordOnly");
                    newPrintHasRecordOnly.SetAttribute("Checked", "True");
                    config.AppendChild(newPrintHasRecordOnly);
                    cd.SetXml("XmlData", config);
                }

                //�C�L�ǥͲM��
                if (PrintStudentList != null)
                {
                    if (PrintStudentList.HasAttribute("Checked"))
                    {
                        _PrintStudentList = bool.Parse(PrintStudentList.GetAttribute("Checked"));
                    }
                }
                else
                {
                    XmlElement newPrintStudentList = config.OwnerDocument.CreateElement("PrintStudentList");
                    newPrintStudentList.SetAttribute("Checked", "False");
                    config.AppendChild(newPrintStudentList);
                    cd.SetXml("XmlData", config);
                }

                if (receive != null)
                {
                    _receiveName = receive.GetAttribute("Name");
                    _receiveAddress = receive.GetAttribute("Address");
                }
                else
                {
                    XmlElement newReceive = config.OwnerDocument.CreateElement("Receive");
                    newReceive.SetAttribute("Name", "");
                    newReceive.SetAttribute("Address", "");
                    config.AppendChild(newReceive);
                    //CurrentUser.Instance.Preference["���m�q����"] = config;
                    cd.SetXml("XmlData", config);
                }

                if (conditions != null)
                {
                    if (conditions.HasAttribute("ConditionName") && conditions.HasAttribute("ConditionNumber"))
                    {
                        _conditionName = conditions.GetAttribute("ConditionName");
                        _conditionNumber = conditions.GetAttribute("ConditionNumber");
                    }
                    else
                    {
                        _conditionName = "";
                        _conditionNumber = "0";
                    }
                }
                else
                {
                    XmlElement newConditions = config.OwnerDocument.CreateElement("Conditions");
                    newConditions.SetAttribute("ConditionName", "");
                    newConditions.SetAttribute("ConditionNumber", "0");
                    config.AppendChild(newConditions);
                    cd.SetXml("XmlData", config);
                    //CurrentUser.Instance.Preference["���g�q����"] = config;
                }

                if (conditions2 != null)
                {
                    if (conditions2.HasAttribute("ConditionName2") && conditions2.HasAttribute("ConditionNumber2"))
                    {
                        _conditionName2 = conditions2.GetAttribute("ConditionName2");
                        _conditionNumber2 = conditions2.GetAttribute("ConditionNumber2");
                    }
                    else
                    {
                        _conditionName2 = "";
                        _conditionNumber2 = "0";
                    }
                }
                else
                {
                    XmlElement newConditions = config.OwnerDocument.CreateElement("Conditions2");
                    newConditions.SetAttribute("ConditionName2", "");
                    newConditions.SetAttribute("ConditionNumber2", "0");
                    config.AppendChild(newConditions);
                    cd.SetXml("XmlData", config);
                    //CurrentUser.Instance.Preference["���g�q����"] = config;
                }

                if (dateRangeMode != null)
                {
                    _mode = AbsenceNotificationConfigForm.DateRangeMode.Custom;
                    
                    //_mode = (AbsenceNotificationConfigForm.DateRangeMode)int.Parse(dateRangeMode.InnerText);

                    if (_mode != AbsenceNotificationConfigForm.DateRangeMode.Custom)
                        dateTimeInput2.Enabled = false;
                    else
                        dateTimeInput2.Enabled = true;
                }
                else
                {
                    XmlElement newDateRangeMode = config.OwnerDocument.CreateElement("DateRangeMode");
                    newDateRangeMode.InnerText = ((int)_mode).ToString();
                    config.AppendChild(newDateRangeMode);
                    //CurrentUser.Instance.Preference["���m�q����"] = config;
                    cd.SetXml("XmlData", config);
                }
            }
            else
            {
                #region ���ͪťճ]�w��
                config = new XmlDocument().CreateElement("���m�q����");
                config.SetAttribute("Default", "true");
                XmlElement printSetup = config.OwnerDocument.CreateElement("PrintHasRecordOnly");
                XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
                XmlElement dateRangeMode = config.OwnerDocument.CreateElement("DateRangeMode");
                XmlElement receive = config.OwnerDocument.CreateElement("Receive");
                XmlElement printStudentList = config.OwnerDocument.CreateElement("PrintStudentList");

                printSetup.SetAttribute("Checked", "true");
                dateRangeMode.InnerText = ((int)_mode).ToString();
                receive.SetAttribute("Name", "");
                receive.SetAttribute("Address", "");
                printStudentList.SetAttribute("Checked", "false");

                config.AppendChild(printSetup);
                config.AppendChild(customize);
                config.AppendChild(dateRangeMode);
                config.AppendChild(receive);
                config.AppendChild(printStudentList);
                //CurrentUser.Instance.Preference["���m�q����"] = config;
                cd.SetXml("XmlData", config);

                _useDefaultTemplate = true;
                _printHasRecordOnly = true;
                _PrintStudentList = false;

                #endregion
            }

            cd.Save();
            #endregion

            _preferenceLoaded = true;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            AbsenceNotificationConfigForm configForm = new AbsenceNotificationConfigForm(
                _useDefaultTemplate, _printHasRecordOnly, _mode, _buffer, _receiveName, _receiveAddress, _conditionName, _conditionNumber, _conditionName2, _conditionNumber2, _PrintStudentList);

            if (configForm.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
                InitialDateRange();
            }
        }

        private void InitialDateRange()
        {
            switch (_mode)
            {
                case AbsenceNotificationConfigForm.DateRangeMode.Month: //��
                    {
                        DateTime a = dateTimeInput1.Value;
                        a = GetMonthFirstDay(a);
                        dateTimeInput1.Text = a.ToShortDateString();
                        dateTimeInput2.Text = a.AddMonths(1).AddDays(-1).ToShortDateString();
                        break;
                    }
                case AbsenceNotificationConfigForm.DateRangeMode.Week: //�g
                    {
                        DateTime b = dateTimeInput1.Value;
                        b = GetWeekFirstDay(b);
                        dateTimeInput1.Text = b.ToShortDateString();
                        dateTimeInput2.Text = b.AddDays(5).ToShortDateString();
                        break;
                    }
                case AbsenceNotificationConfigForm.DateRangeMode.Custom: //�ۭq
                    {
                        //dateTimeInput2.Text = dateTimeInput1.Text = DateTime.Today.ToShortDateString();
                        break;
                    }
                default:
                    throw new Exception("Date Range Mode Error.");
            }

            _printable = true;
            _startTextBoxOK = true;
            _endTextBoxOK = true;
        }

        private DateTime GetWeekFirstDay(DateTime inputDate)
        {
            switch (inputDate.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    return inputDate;
                case DayOfWeek.Tuesday:
                    return inputDate.AddDays(-1);
                case DayOfWeek.Wednesday:
                    return inputDate.AddDays(-2);
                case DayOfWeek.Thursday:
                    return inputDate.AddDays(-3);
                case DayOfWeek.Friday:
                    return inputDate.AddDays(-4);
                case DayOfWeek.Saturday:
                    return inputDate.AddDays(-5);
                default:
                    return inputDate.AddDays(-6);
            }
        }

        private DateTime GetMonthFirstDay(DateTime inputDate)
        {
            return DateTime.Parse(inputDate.Year + "/" + inputDate.Month + "/1");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (_printable)
            //    dateTimeInput1.Text = _startDate.ToShortDateString();
            //timer1.Stop();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm form = new SelectTypeForm(addconfigName);
            form.ShowDialog();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            if (_startTextBoxOK && _mode != AbsenceNotificationConfigForm.DateRangeMode.Custom)
            {
                switch (_mode)
                {
                    case AbsenceNotificationConfigForm.DateRangeMode.Month: //��
                        {
                            _startDate = GetMonthFirstDay(DateTime.Parse(dateTimeInput1.Text));
                            _endDate = _startDate.AddMonths(1).AddDays(-1);
                            dateTimeInput1.Text = _startDate.ToShortDateString();
                            dateTimeInput2.Text = _endDate.ToShortDateString();
                            _printable = true;
                            break;
                        }
                    case AbsenceNotificationConfigForm.DateRangeMode.Week: //�g
                        {
                            _startDate = GetWeekFirstDay(DateTime.Parse(dateTimeInput1.Text));
                            _endDate = _startDate.AddDays(4);
                            dateTimeInput1.Text = _startDate.ToShortDateString();
                            dateTimeInput2.Text = _endDate.ToShortDateString();
                            _printable = true;
                            break;
                        }
                    case AbsenceNotificationConfigForm.DateRangeMode.Custom: //�ۭq
                        break;
                    default:
                        throw new Exception("Date Range Mode Error");
                }

                //if (dateTimeInput1.Text != _startDate.ToShortDateString() && timer1 != null)
                //    timer1.Start();
                errorProvider1.Clear();
            }
        }

        private void dateTimeInput2_TextChanged(object sender, EventArgs e)
        {
            //if (_preferenceLoaded)
            //{
            //    if (_mode == DateRangeMode.Custom)
            //    {
            //        base.textBoxX2_TextChanged(sender, e);
            //    }
            //    else
            //    {
            //        _endTextBoxOK = true;
            //        errorProvider2.Clear();
            //    }
            //}
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "���m�q����_�\���ܼ��`��.docx";
            sfd.Filter = "Word�ɮ� (*.docx)|*.docx|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.���m�q����_�\���ܼ��`��, 0, Properties.Resources.���m�q����_�\���ܼ��`��.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void cbUpdate_CheckedChanged(object sender, EventArgs e)
        {
            _cbPaper = cbUpdate.Checked;
        }
    }
}