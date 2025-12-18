using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace DEMO
{
    public partial class ucTimeOfHour : DateTimePicker
    {
        protected DateTime DefauleTime
        {
            get { return (new DateTime(2000, 1, 1, 0, 0, 0)); }
        }

        public int UI_Minute
        {
            get { return (base.Value.Minute); }
            set 
            {
                int tempMinute = value;
                // 防呆
                if (tempMinute > 59)
                    tempMinute = 59;
                else if (tempMinute < 0)
                    tempMinute = 0;

                base.Value = this.DefauleTime.AddMinutes(tempMinute).AddSeconds(base.Value.Second);
            }
        }

        public int UI_Second
        {
            get { return (base.Value.Second); }
            set 
            {
                int tempSecond = value;
                // 防呆
                if (tempSecond > 59)
                    tempSecond = 59;
                else if (tempSecond < 0)
                    tempSecond = 0;

                // 設定 新的時間值
                base.Value = this.DefauleTime.AddMinutes(base.Value.Minute).AddSeconds(tempSecond); 
            }
        }

        public TimeSpan UI_TimeSpan
        {
            get
            {
                DateTime tempDT = new DateTime(base.Value.Year, base.Value.Month, base.Value.Day, 0, 0, 0);
                return (base.Value - tempDT);
            }
            set { base.Value = this.DefauleTime.AddMinutes(value.Minutes).AddSeconds(value.Seconds); }
        }

        public ucTimeOfHour()
        {
            InitializeComponent();

            this.Format = DateTimePickerFormat.Custom;
            this.CustomFormat = "mm:ss";
            this.Size = new Size(60, 22);
            this.ShowUpDown = true;
            base.Value = this.DefauleTime;
            base.Font = new Font(SystemFonts.DefaultFont.FontFamily.Name, 10, base.Font.Style, base.Font.Unit, base.Font.GdiCharSet, base.Font.GdiVerticalFont);
        }

    }
}
