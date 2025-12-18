namespace DEMO
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.clStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button2 = new System.Windows.Forms.Button();
            this.clResult = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.gvwRAC2000 = new System.Windows.Forms.DataGridView();
            this.clAddress = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clPort = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gvwRAC2000)).BeginInit();
            this.SuspendLayout();
            // 
            // clStatus
            // 
            this.clStatus.HeaderText = "STATUS";
            this.clStatus.Name = "clStatus";
            // 
            // clId
            // 
            this.clId.HeaderText = "RAC2000ID";
            this.clId.Name = "clId";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(493, 275);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(140, 23);
            this.button2.TabIndex = 5;
            this.button2.Text = "UNCONNECT";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // clResult
            // 
            this.clResult.HeaderText = "RESULT CONTENT";
            this.clResult.Name = "clResult";
            this.clResult.Width = 200;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(347, 275);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(140, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "CONNECT";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // gvwRAC2000
            // 
            this.gvwRAC2000.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvwRAC2000.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clAddress,
            this.clPort,
            this.clId,
            this.clStatus,
            this.clResult});
            this.gvwRAC2000.Dock = System.Windows.Forms.DockStyle.Top;
            this.gvwRAC2000.Location = new System.Drawing.Point(0, 0);
            this.gvwRAC2000.Name = "gvwRAC2000";
            this.gvwRAC2000.RowTemplate.Height = 23;
            this.gvwRAC2000.Size = new System.Drawing.Size(644, 256);
            this.gvwRAC2000.TabIndex = 3;
            this.gvwRAC2000.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.gvwRAC2000_DefaultValuesNeeded);
            // 
            // clAddress
            // 
            this.clAddress.HeaderText = "TCP IP/COM";
            this.clAddress.Name = "clAddress";
            // 
            // clPort
            // 
            this.clPort.HeaderText = "PORT/RATE";
            this.clPort.Name = "clPort";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(644, 320);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.gvwRAC2000);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SDK Demo";
            ((System.ComponentModel.ISupportInitialize)(this.gvwRAC2000)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridViewTextBoxColumn clStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn clId;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridViewTextBoxColumn clResult;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView gvwRAC2000;
        private System.Windows.Forms.DataGridViewTextBoxColumn clAddress;
        private System.Windows.Forms.DataGridViewTextBoxColumn clPort;
    }
}

