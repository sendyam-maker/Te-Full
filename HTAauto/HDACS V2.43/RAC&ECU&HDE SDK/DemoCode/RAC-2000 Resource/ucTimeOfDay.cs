using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace DEMO
{
    public partial class ucTimeOfDay : DateTimePicker
    {
        protected DateTime DefauleTime
        {
            get { return (new DateTime(2000, 1, 1, 0, 0, 0)); }
        }

        public int UI_Hour
        {
            get { return (base.Value.Hour); }
            set
            {
                int tempHour = value;
                // 防呆
                if (tempHour > 23)
                    tempHour = 23;
                else if (tempHour < 0)
                    tempHour = 0;

                // 設定 新的時間值
                base.Value = this.DefauleTime.AddHours(tempHour).AddMinutes(base.Value.Minute);
            }
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

                base.Value = this.DefauleTime.AddHours(base.Value.Hour).AddMinutes(tempMinute);
            }
        }

        public TimeSpan UI_TimeSpan
        {
            get
            {
                DateTime tempDT = new DateTime(base.Value.Year, base.Value.Month, base.Value.Day, 0, 0, 0);
                return (base.Value - tempDT);
            }
            set { base.Value = this.DefauleTime.AddHours(value.Hours).AddMinutes(value.Minutes); }
        }

        public ucTimeOfDay()
        {
            InitializeComponent();
            this.Format = DateTimePickerFormat.Custom;
            this.CustomFormat = "HH:mm";
            this.Size = new Size(60, 22);
            this.ShowUpDown = true;
            base.Value = this.DefauleTime;
            base.Font = new Font(SystemFonts.DefaultFont.FontFamily.Name, 10, base.Font.Style, base.Font.Unit, base.Font.GdiCharSet, base.Font.GdiVerticalFont);
        }

        public ucTimeOfDay(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
    }
}
