using System;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace hwhs.epost.學期成績通知單
{
    public partial class SelectDateRangeForm : BaseForm
    {
        protected DateTime _startDate;
        protected DateTime _endDate;
        protected bool _startTextBoxOK = false;
        protected bool _endTextBoxOK = false;
        protected bool _printable = false;

        public DateTime StartDate
        {
            get { return _startDate; }
        }

        public DateTime EndDate
        {
            get { return _endDate; }
        }

        public SelectDateRangeForm(string title) : this()
        {
            Text = title;
        }

        public SelectDateRangeForm()
        {
            InitializeComponent();
            _startDate = DateTime.Today;
            _endDate = DateTime.Today;
            dateTimeInput1.Value = _startDate;
            dateTimeInput2.Value = _endDate;
        }

        protected virtual void buttonX1_Click(object sender, EventArgs e)
        {
            if (_printable == true)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private bool ValidateRange(string startDate, string endDate)
        {
            DateTime a, b;
            a = DateTime.Parse(startDate);
            b = DateTime.Parse(endDate);

            if (DateTime.Compare(b, a) < 0)
            {
                _printable = false;
                return false;
            }
            else
            {
                _printable = true;
                _startDate = a;
                _endDate = b;
                return true;
            }
        }

        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            _startTextBoxOK = true;
            if (_endTextBoxOK)
            {
                if (!ValidateRange(dateTimeInput1.Text, dateTimeInput2.Text))
                    errorProvider1.SetError(dateTimeInput1, "日期區間錯誤");
                else
                {
                    errorProvider1.Clear();
                    errorProvider2.Clear();
                }
            }
        }

        private void dateTimeInput2_TextChanged(object sender, EventArgs e)
        {
            errorProvider2.Clear();
            _endTextBoxOK = true;
            if (_startTextBoxOK)
            {
                if (!ValidateRange(dateTimeInput1.Text, dateTimeInput2.Text))
                    errorProvider2.SetError(dateTimeInput2, "日期區間錯誤");
                else
                {
                    errorProvider1.Clear();
                    errorProvider2.Clear();
                }
            }

        }
    }
}