namespace Demo
{
    partial class FrmMain
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
            this.components = new System.ComponentModel.Container();
            this.panel1 = new System.Windows.Forms.Panel();
            this.gvwHTAList = new System.Windows.Forms.DataGridView();
            this.clAddress = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clPort = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clDeviceId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clResult = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.connectHTAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.unconnectHTAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvwHTAList)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.gvwHTAList);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(766, 388);
            this.panel1.TabIndex = 0;
            // 
            // gvwHTAList
            // 
            this.gvwHTAList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvwHTAList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clAddress,
            this.clPort,
            this.clDeviceId,
            this.clStatus,
            this.clResult});
            this.gvwHTAList.ContextMenuStrip = this.contextMenuStrip1;
            this.gvwHTAList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gvwHTAList.Location = new System.Drawing.Point(0, 0);
            this.gvwHTAList.Name = "gvwHTAList";
            this.gvwHTAList.RowTemplate.Height = 23;
            this.gvwHTAList.Size = new System.Drawing.Size(766, 388);
            this.gvwHTAList.TabIndex = 0;
            this.gvwHTAList.UserAddedRow += new System.Windows.Forms.DataGridViewRowEventHandler(this.gvwHTAList_UserAddedRow);
            // 
            // clAddress
            // 
            this.clAddress.HeaderText = "TCP/COM";
            this.clAddress.Name = "clAddress";
            // 
            // clPort
            // 
            this.clPort.HeaderText = "PORT/RATE";
            this.clPort.Name = "clPort";
            // 
            // clDeviceId
            // 
            this.clDeviceId.HeaderText = "ID";
            this.clDeviceId.Name = "clDeviceId";
            // 
            // clStatus
            // 
            this.clStatus.HeaderText = "STATUS";
            this.clStatus.Name = "clStatus";
            this.clStatus.ReadOnly = true;
            // 
            // clResult
            // 
            this.clResult.HeaderText = "RETURN CONTENT";
            this.clResult.Name = "clResult";
            this.clResult.ReadOnly = true;
            this.clResult.Width = 300;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.connectHTAToolStripMenuItem,
            this.unconnectHTAToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(147, 48);
            // 
            // connectHTAToolStripMenuItem
            // 
            this.connectHTAToolStripMenuItem.Name = "connectHTAToolStripMenuItem";
            this.connectHTAToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            this.connectHTAToolStripMenuItem.Text = "Connect HTA";
            this.connectHTAToolStripMenuItem.Click += new System.EventHandler(this.connectHTAToolStripMenuItem_Click);
            // 
            // unconnectHTAToolStripMenuItem
            // 
            this.unconnectHTAToolStripMenuItem.Name = "unconnectHTAToolStripMenuItem";
            this.unconnectHTAToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            this.unconnectHTAToolStripMenuItem.Text = "Unconnect HTA";
            this.unconnectHTAToolStripMenuItem.Click += new System.EventHandler(this.unconnectHTAToolStripMenuItem_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.button2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 288);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(766, 100);
            this.panel2.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(247, 27);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(103, 33);
            this.button1.TabIndex = 2;
            this.button1.Text = "Unconnect HTA";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(83, 27);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(103, 33);
            this.button2.TabIndex = 1;
            this.button2.Text = "Connect HTA";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 88);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(11, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "_";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(766, 388);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MainUI";
            this.Load += new System.EventHandler(this.FrmMain_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gvwHTAList)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView gvwHTAList;
        private System.Windows.Forms.DataGridViewTextBoxColumn clAddress;
        private System.Windows.Forms.DataGridViewTextBoxColumn clPort;
        private System.Windows.Forms.DataGridViewTextBoxColumn clDeviceId;
        private System.Windows.Forms.DataGridViewTextBoxColumn clStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn clResult;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem connectHTAToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem unconnectHTAToolStripMenuItem;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
    }
}