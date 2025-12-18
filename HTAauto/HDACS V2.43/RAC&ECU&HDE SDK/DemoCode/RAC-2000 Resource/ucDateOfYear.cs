using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace DEMO
{
    public partial class ucDateOfYear : DateTimePicker
    {
        protected DateTime DefauleTime
        {
            get { return (new DateTime(2000, 1, 1, 0, 0, 0)); }
        }

        protected void SetValue(int inMonth, int inDay)
        {
            int tempMonth = inMonth;
            // 防呆
            if (tempMonth > 12)
                tempMonth = 12;
            else if (tempMonth < 1)
                tempMonth = 1;

            int tempMaxDay = 29;
            if (tempMonth == 1 || tempMonth == 3 || tempMonth == 5 || tempMonth == 7 || tempMonth == 8 || tempMonth == 10 || tempMonth == 12)
                tempMaxDay = 31;
            else if (tempMonth == 4 || tempMonth == 6 || tempMonth == 9 || tempMonth == 11)
                tempMaxDay = 30;

            int tempDay = inDay;
            // 防呆
            if (tempDay > tempMaxDay)
                tempDay = tempMaxDay;
            else if (tempDay < 1)
                tempDay = 1;

            // 設定 新的日期值
            base.Value = new DateTime(DefauleTime.Year, tempMonth, tempDay, DefauleTime.Hour, DefauleTime.Minute, DefauleTime.Second);
        }

        public int UI_Month
        {
            get { return (base.Value.Month); }
            set { this.SetValue(value, base.Value.Day); }
        }

        public int UI_Day
        {
            get { return (base.Value.Day); }
            set { this.SetValue(base.Value.Month, value); }
        }

        public ucDateOfYear()
        {
            InitializeComponent();

            this.Format = DateTimePickerFormat.Custom;
            this.CustomFormat = "MM/dd";
            this.Size = new Size(60, 22);
            this.ShowUpDown = true;
            base.Value = this.DefauleTime;
            base.Font = new Font(SystemFonts.DefaultFont.FontFamily.Name, 10, base.Font.Style, base.Font.Unit, base.Font.GdiCharSet, base.Font.GdiVerticalFont);
        }

    }
}
