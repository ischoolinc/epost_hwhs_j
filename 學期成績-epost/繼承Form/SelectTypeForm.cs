using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Xml;
using FISCA.DSAUtil;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;
using K12.Data;

namespace hwhs.epost.學期成績通知單
{
    public partial class SelectTypeForm : BaseForm
    {
        private string _preferenceElementName;
        private BackgroundWorker _BGWAbsenceAndPeriodList;

        private List<string> typeList = new List<string>();
        private List<string> absenceList = new List<string>();

        bool valueOnChange=false;

        public SelectTypeForm(string name)
        {
            InitializeComponent();

            _preferenceElementName = name;

            _BGWAbsenceAndPeriodList = new BackgroundWorker();
            _BGWAbsenceAndPeriodList.DoWork += new DoWorkEventHandler(_BGWAbsenceAndPeriodList_DoWork);
            _BGWAbsenceAndPeriodList.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWAbsenceAndPeriodList_RunWorkerCompleted);
            _BGWAbsenceAndPeriodList.RunWorkerAsync();
        }

        void _BGWAbsenceAndPeriodList_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            System.Windows.Forms.DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn();
            colName.HeaderText = "節次分類";
            colName.MinimumWidth = 70;
            colName.Name = "colName";
            colName.ReadOnly = true;
            colName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            colName.Width = 70;
            this.dataGridViewX1.Columns.Add(colName);

            foreach (string absence in absenceList)
            {
                System.Windows.Forms.DataGridViewCheckBoxColumn newCol=new DataGridViewCheckBoxColumn();
                newCol.HeaderText = absence;
                newCol.Width = 55;
                newCol.ReadOnly = false;
                newCol.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
                newCol.Tag = absence ;
                newCol.ValueType=typeof(bool);
                this.dataGridViewX1.Columns.Add(newCol);
            }
            foreach (string type in typeList)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1, type);
                row.Tag = type;
                dataGridViewX1.Rows.Add(row);
            }

            #region 讀取列印設定 Preference
            valueOnChange = true;
            //XmlElement config = CurrentUser.Instance.Preference[_preferenceElementName];
            ConfigData cd = School.Configuration[_preferenceElementName];
            string strr = cd["XmlData"];
            XmlElement config;

            if (strr != "")
            {
                config = DSXmlHelper.LoadXml(strr);
                #region 已有設定檔則將設定檔內容填回畫面上
                foreach (XmlElement type in config.SelectNodes("Type"))
                {
                    string typeName = type.GetAttribute("Text");
                    foreach (DataGridViewRow row in dataGridViewX1.Rows)
                    {
                        if (typeName == ("" + row.Tag))
                        {
                            foreach (XmlElement absence in type.SelectNodes("Absence"))
                            {
                                string absenceName = absence.GetAttribute("Text");
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    if (cell.OwningColumn is DataGridViewCheckBoxColumn && ("" + cell.OwningColumn.Tag) == absenceName)
                                    {
                                        cell.Value = true;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
                #endregion
            }
            else
            {
                #region 產生空白設定檔
                config = new XmlDocument().CreateElement(_preferenceElementName);
                //CurrentUser.Instance.Preference[_preferenceElementName] = config;
                cd.SetXml("XmlData", config);
                #endregion
            }

            cd.Save();

            valueOnChange = false;

            #endregion
        }

        void _BGWAbsenceAndPeriodList_DoWork(object sender, DoWorkEventArgs e)
        {
            foreach (PeriodMappingInfo each in K12.Data.PeriodMapping.SelectAll())
            {
                if (!typeList.Contains(each.Type))
                    typeList.Add(each.Type);
            }

            foreach (AbsenceMappingInfo each in K12.Data.AbsenceMapping.SelectAll())
            {
                if (!absenceList.Contains(each.Name))
                    absenceList.Add(each.Name);
            }
        }

        //private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        //{
        //    TreeNode checkedNode = e.Node;

        //    if (typeList.Contains(checkedNode.Text) && checkedNode.Parent == null)
        //    {
        //        foreach (TreeNode subnode in checkedNode.Nodes)
        //        {
        //            subnode.Checked = checkedNode.Checked;
        //        }
        //    }
        //}

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (!CheckColumnNumber())
                return;

            #region 更新列印設定 Preference

            //XmlElement config = CurrentUser.Instance.Preference[_preferenceElementName];
            ConfigData cd = School.Configuration[_preferenceElementName];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config == null)
            {
                config = new XmlDocument().CreateElement(_preferenceElementName);
            }

            config.RemoveAll();

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                bool needToAppend = false;
                XmlElement type = config.OwnerDocument.CreateElement("Type");
                type.SetAttribute("Text", "" + row.Tag);
                foreach (DataGridViewCell cell in row.Cells)
                {
                    XmlElement absence = config.OwnerDocument.CreateElement("Absence");
                    absence.SetAttribute("Text", ""+cell.OwningColumn.Tag);
                    if (cell.Value is bool && ((bool)cell.Value))
                    {
                        needToAppend = true;
                        type.AppendChild(absence);
                    }
                }
                if(needToAppend)
                    config.AppendChild(type);
            }

            //foreach (TreeNode typeNode in treeView1.Nodes)
            //{
            //    XmlElement type = config.OwnerDocument.CreateElement("Type");
            //    type.SetAttribute("Text", typeNode.Text);
            //    type.SetAttribute("Checked", typeNode.Checked.ToString());

            //    foreach (TreeNode absenceNode in typeNode.Nodes)
            //    {
            //        if (absenceNode.Checked == true)
            //        {
            //            XmlElement absence = config.OwnerDocument.CreateElement("Absence");
            //            absence.SetAttribute("Text", absenceNode.Text);
            //            type.AppendChild(absence);
            //        }
            //    }
            //    config.AppendChild(type);
            //}


            //CurrentUser.Instance.Preference[_preferenceElementName] = config;
            cd.SetXml("XmlData", config);
            cd.Save();

            #endregion

            this.Close();
        }

        internal bool CheckColumnNumber()
        {
            int limit = 253;
            int columnNumber = 0;
            int block = 9;

            //foreach (TreeNode type in treeView1.Nodes)
            //{
            //    foreach (TreeNode var in type.Nodes)
            //    {
            //        if (var.Checked == true)
            //            columnNumber++;
            //    }
            //}
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value is bool &&((bool)cell.Value ))
                        columnNumber++;
                }
            }

            if (columnNumber * block > limit)
            {
                MsgBox.Show("您所選擇的假別超出 Excel 的最大欄位，請減少部分假別");
                return false;
            }
            else
                return true;
        }

        private void dataGridViewX1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow each1 in dataGridViewX1.Rows)
            {
                foreach(DataGridViewCell each2 in each1.Cells)
                {
                    if (each2.Value is string)
                        continue;

                    each2.Value = checkBoxX1.Checked;
                }
            }
        }

        //private void dataGridViewX1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        //{
        //    if (e.ColumnIndex < 0 || e.RowIndex < 0)
        //        return;
        //    if ( valueOnChange)
        //        return;
        //    else
        //        valueOnChange = true;
        //    DataGridViewCell checkedCell=dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];
        //    foreach (DataGridViewCell  cell in dataGridViewX1.SelectedCells)
        //    {
        //        if (cell.OwningColumn is DataGridViewCheckBoxColumn && cell != checkedCell)
        //            cell.Value = checkedCell.Value;
        //    }
        //    valueOnChange = false;
        //}
    }
}