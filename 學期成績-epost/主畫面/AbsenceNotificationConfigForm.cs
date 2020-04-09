using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar.Rendering;
using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;

namespace hwhs.epost.定期評量通知單
{
    public partial class AbsenceNotificationConfigForm : BaseForm
    {
        private byte[] _buffer = null;
        private string base64 = null;
        private bool _isUpload = false;
        private bool _defaultTemplate;
        private bool _printHasRecordOnly;
        private DateRangeMode _mode = DateRangeMode.Month;
        private bool _printStudentList;

        string configName = "定期評量通知單_2019_弘文epost";
        string addconfigName = "定期評量通知單_缺曠別設定_2019_弘文epost";

        public AbsenceNotificationConfigForm(bool defaultTemplate, bool printHasRecordOnly, DateRangeMode mode, byte[] buffer, string name, string address, string conditionName, string conditionNumber, string conditionName2, string conditionNumber2, bool printStudentList)
        {
            InitializeComponent();
            #region 如果系統的Renderer是Office2007Renderer，同化_ClassTeacherView,_CategoryView的顏色
            if (GlobalManager.Renderer is Office2007Renderer)
            {
                ((Office2007Renderer)GlobalManager.Renderer).ColorTableChanged += new EventHandler(ScoreCalcRuleEditor_ColorTableChanged);
                SetForeColor(this);
            }
            #endregion
            _defaultTemplate = defaultTemplate;
            _printHasRecordOnly = printHasRecordOnly;
            _mode = mode;
            _printStudentList = printStudentList;

            if (buffer != null)
                _buffer = buffer;

            if (defaultTemplate)
                radioButton1.Checked = true;
            else
                radioButton2.Checked = true;

            checkBoxX1.Checked = printHasRecordOnly;
            checkBoxX2.Checked = printStudentList;

            switch (mode)
            {
                case DateRangeMode.Month:
                    radioButton3.Checked = true;
                    break;
                case DateRangeMode.Week:
                    radioButton4.Checked = true;
                    break;
                case DateRangeMode.Custom:
                    radioButton5.Checked = true;
                    break;
                default:
                    throw new Exception("Date Range Mode Error.");
            }

            comboBoxEx1.SelectedIndex = 0;
            comboBoxEx2.SelectedIndex = 0;

            foreach (DevComponents.Editors.ComboItem var in comboBoxEx1.Items)
            {
                if (var.Text == name)
                {
                    comboBoxEx1.SelectedIndex = comboBoxEx1.Items.IndexOf(var);
                    break;
                }
            }

            foreach (DevComponents.Editors.ComboItem var in comboBoxEx2.Items)
            {
                if (var.Text == address)
                {
                    comboBoxEx2.SelectedIndex = comboBoxEx2.Items.IndexOf(var);
                    break;
                }
            }
            decimal tryValue;
            numericUpDown1.Value = (decimal.TryParse(conditionNumber, out tryValue)) ? tryValue : 0;

            numericUpDown2.Value = (decimal.TryParse(conditionNumber2, out tryValue)) ? tryValue : 0;

            GetAbsenceConfig(); //取得缺曠別

            foreach (string each in comboBoxEx3.Items) //將畫面設定為前次設定值
            {
                if (each == conditionName)
                {
                    comboBoxEx3.SelectedItem = each;
                }
            }
            foreach (string each in comboBoxEx4.Items) //將畫面設定為前次設定值
            {
                if (each == conditionName2)
                {
                    comboBoxEx4.SelectedItem = each;
                }
            }
        }

        private void GetAbsenceConfig()
        {
            #region 取得使用者自己設定的內容
            List<string> list = new List<string>();
            list.Add("");
            ConfigData cd = K12.Data.School.Configuration[addconfigName];
            string strr = cd["XmlData"];

            if (strr != "")
            {
                XmlElement Config = DSXmlHelper.LoadXml(strr);

                foreach (XmlElement each in Config.SelectNodes("Type"))
                {
                    foreach (XmlElement eachXX in each.SelectNodes("Absence"))
                    {
                        if (!list.Contains(eachXX.GetAttribute("Text"))) //如果假別不存在於清單
                        {
                            list.Add(eachXX.GetAttribute("Text"));
                        }
                    }
                }
            }

            foreach (string each in list)
            {
                comboBoxEx3.Items.Add(each);
                comboBoxEx4.Items.Add(each);
            }
            #endregion
        }

        void ScoreCalcRuleEditor_ColorTableChanged(object sender, EventArgs e)
        {
            SetForeColor(this);
        }

        private void SetForeColor(Control parent)
        {
            foreach (Control var in parent.Controls)
            {
                if (var is RadioButton)
                    var.ForeColor = ((Office2007Renderer)GlobalManager.Renderer).ColorTable.CheckBoxItem.Default.Text;
                SetForeColor(var);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                radioButton2.Checked = false;
                _defaultTemplate = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                radioButton1.Checked = false;
                _defaultTemplate = false;
            }
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            _printHasRecordOnly = checkBoxX1.Checked;
        }

        private void checkBoxX2_CheckedChanged(object sender, EventArgs e)
        {
            _printStudentList = checkBoxX2.Checked;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "缺曠通知單.docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.缺曠通知單_住址中間版, 0, Properties.Resources.缺曠通知單_住址中間版.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "自訂缺曠通知單.docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(new MemoryStream(_buffer));
                    doc.Save(sfd.FileName, Aspose.Words.SaveFormat.Docx);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案無法儲存。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("檔案無法開啟。" + ex.Message);
                    return;
                }
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇自訂的缺曠通知單範本";
            ofd.Filter = "Word檔案 (*.docx)|*.docx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                    byte[] tempBuffer = new byte[fs.Length];
                    fs.Read(tempBuffer, 0, tempBuffer.Length);
                    base64 = Convert.ToBase64String(tempBuffer);
                    _isUpload = true;
                    fs.Close();
                    MsgBox.Show("上傳成功。");
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference

            //XmlElement config = CurrentUser.Instance.Preference["缺曠通知單"];
            ConfigData cd = K12.Data.School.Configuration[configName];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config == null)
            {
                config = new XmlDocument().CreateElement("定期評量通知單");
            }

            config.SetAttribute("Default", _defaultTemplate.ToString());

            XmlElement printSetup = config.OwnerDocument.CreateElement("PrintHasRecordOnly");
            XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
            XmlElement mode = config.OwnerDocument.CreateElement("DateRangeMode");
            XmlElement receive = config.OwnerDocument.CreateElement("Receive");
            XmlElement conditions = config.OwnerDocument.CreateElement("Conditions");
            XmlElement conditions2 = config.OwnerDocument.CreateElement("Conditions2");
            XmlElement PrintStudentList = config.OwnerDocument.CreateElement("PrintStudentList");

            printSetup.SetAttribute("Checked", _printHasRecordOnly.ToString());
            PrintStudentList.SetAttribute("Checked", _printStudentList.ToString());

            config.ReplaceChild(printSetup, config.SelectSingleNode("PrintHasRecordOnly"));
            config.ReplaceChild(PrintStudentList, config.SelectSingleNode("PrintStudentList"));

            if (_isUpload)
            {
                customize.InnerText = base64;
                config.ReplaceChild(customize, config.SelectSingleNode("CustomizeTemplate"));
            }

            mode.InnerText = ((int)_mode).ToString();
            config.ReplaceChild(mode, config.SelectSingleNode("DateRangeMode"));


            receive.SetAttribute("Name", ((DevComponents.Editors.ComboItem)comboBoxEx1.SelectedItem).Text);
            receive.SetAttribute("Address", ((DevComponents.Editors.ComboItem)comboBoxEx2.SelectedItem).Text);
            if (config.SelectSingleNode("Receive") == null)
                config.AppendChild(receive);
            else
                config.ReplaceChild(receive, config.SelectSingleNode("Receive"));

            #region ConditionName

            //設定Attribute,為comboBoxEx3所選
            conditions.SetAttribute("ConditionName", ((string)comboBoxEx3.SelectedItem));
            //設定Attribute,為numericUpDown2內容
            conditions.SetAttribute("ConditionNumber", numericUpDown1.Value.ToString());

            //如果config設定值為null
            if (config.SelectSingleNode("Conditions") == null)
                config.AppendChild(conditions); //加入Child
            else
                config.ReplaceChild(conditions, config.SelectSingleNode("Conditions")); //否則取代掉Child

            #endregion

            #region ConditionName2

            //設定Attribute,為comboBoxEx4所選
            conditions2.SetAttribute("ConditionName2", ((string)comboBoxEx4.SelectedItem));
            //設定Attribute,為numericUpDown2內容
            conditions2.SetAttribute("ConditionNumber2", numericUpDown2.Value.ToString());

            //如果config設定值為null
            if (config.SelectSingleNode("Conditions2") == null)
                config.AppendChild(conditions2); //加入Child
            else
                config.ReplaceChild(conditions2, config.SelectSingleNode("Conditions2")); //否則取代掉Child

            #endregion

            //CurrentUser.Instance.Preference["缺曠通知單"] = config;
            cd.SetXml("XmlData", config);
            cd.Save();

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                radioButton4.Checked = false;
                radioButton5.Checked = false;
                _mode = DateRangeMode.Month;
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                radioButton3.Checked = false;
                radioButton5.Checked = false;
                _mode = DateRangeMode.Week;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                radioButton3.Checked = false;
                radioButton4.Checked = false;
                _mode = DateRangeMode.Custom;
            }
        }

        public enum DateRangeMode { Month, Week, Custom }
    }
}