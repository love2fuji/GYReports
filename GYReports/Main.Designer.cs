namespace GYReports
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tabConServerLog = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.ServerLog = new System.Windows.Forms.RichTextBox();
            this.tpgRegion = new System.Windows.Forms.TabPage();
            this.dgvShowRegion = new System.Windows.Forms.DataGridView();
            this.tpgDepart = new System.Windows.Forms.TabPage();
            this.dgvShowDepart = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnBaseDataExport2Excel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpStopTime = new System.Windows.Forms.DateTimePicker();
            this.dtpStartTime = new System.Windows.Forms.DateTimePicker();
            this.groupBox2.SuspendLayout();
            this.tabConServerLog.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tpgRegion.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvShowRegion)).BeginInit();
            this.tpgDepart.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvShowDepart)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tabConServerLog);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 75);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(956, 535);
            this.groupBox2.TabIndex = 35;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "状态显示";
            // 
            // tabConServerLog
            // 
            this.tabConServerLog.Controls.Add(this.tabPage1);
            this.tabConServerLog.Controls.Add(this.tpgRegion);
            this.tabConServerLog.Controls.Add(this.tpgDepart);
            this.tabConServerLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabConServerLog.Location = new System.Drawing.Point(3, 17);
            this.tabConServerLog.Name = "tabConServerLog";
            this.tabConServerLog.SelectedIndex = 0;
            this.tabConServerLog.Size = new System.Drawing.Size(950, 515);
            this.tabConServerLog.TabIndex = 28;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.ServerLog);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(942, 489);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "服务日志";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // ServerLog
            // 
            this.ServerLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ServerLog.Location = new System.Drawing.Point(3, 3);
            this.ServerLog.Name = "ServerLog";
            this.ServerLog.Size = new System.Drawing.Size(936, 483);
            this.ServerLog.TabIndex = 24;
            this.ServerLog.Text = "";
            // 
            // tpgRegion
            // 
            this.tpgRegion.AutoScroll = true;
            this.tpgRegion.Controls.Add(this.dgvShowRegion);
            this.tpgRegion.Location = new System.Drawing.Point(4, 22);
            this.tpgRegion.Name = "tpgRegion";
            this.tpgRegion.Padding = new System.Windows.Forms.Padding(3);
            this.tpgRegion.Size = new System.Drawing.Size(942, 489);
            this.tpgRegion.TabIndex = 1;
            this.tpgRegion.Text = "区域/部门基础数据";
            this.tpgRegion.UseVisualStyleBackColor = true;
            // 
            // dgvShowRegion
            // 
            this.dgvShowRegion.AllowDrop = true;
            this.dgvShowRegion.AllowUserToAddRows = false;
            this.dgvShowRegion.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvShowRegion.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvShowRegion.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvShowRegion.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvShowRegion.Location = new System.Drawing.Point(3, 3);
            this.dgvShowRegion.Name = "dgvShowRegion";
            this.dgvShowRegion.RowHeadersWidth = 50;
            this.dgvShowRegion.RowTemplate.Height = 23;
            this.dgvShowRegion.Size = new System.Drawing.Size(936, 483);
            this.dgvShowRegion.TabIndex = 0;
            // 
            // tpgDepart
            // 
            this.tpgDepart.Controls.Add(this.dgvShowDepart);
            this.tpgDepart.Location = new System.Drawing.Point(4, 22);
            this.tpgDepart.Name = "tpgDepart";
            this.tpgDepart.Padding = new System.Windows.Forms.Padding(3);
            this.tpgDepart.Size = new System.Drawing.Size(1160, 487);
            this.tpgDepart.TabIndex = 2;
            this.tpgDepart.Text = " ";
            this.tpgDepart.UseVisualStyleBackColor = true;
            // 
            // dgvShowDepart
            // 
            this.dgvShowDepart.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dgvShowDepart.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvShowDepart.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvShowDepart.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvShowDepart.Location = new System.Drawing.Point(3, 3);
            this.dgvShowDepart.Name = "dgvShowDepart";
            this.dgvShowDepart.RowHeadersWidth = 10;
            this.dgvShowDepart.RowTemplate.Height = 23;
            this.dgvShowDepart.Size = new System.Drawing.Size(1154, 481);
            this.dgvShowDepart.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.dtpStopTime);
            this.groupBox1.Controls.Add(this.dtpStartTime);
            this.groupBox1.Controls.Add(this.btnBaseDataExport2Excel);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(956, 75);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "选项";
            // 
            // btnBaseDataExport2Excel
            // 
            this.btnBaseDataExport2Excel.BackColor = System.Drawing.SystemColors.Control;
            this.btnBaseDataExport2Excel.Location = new System.Drawing.Point(403, 17);
            this.btnBaseDataExport2Excel.Name = "btnBaseDataExport2Excel";
            this.btnBaseDataExport2Excel.Size = new System.Drawing.Size(90, 41);
            this.btnBaseDataExport2Excel.TabIndex = 19;
            this.btnBaseDataExport2Excel.Text = "生成报表";
            this.btnBaseDataExport2Excel.UseVisualStyleBackColor = false;
            this.btnBaseDataExport2Excel.Click += new System.EventHandler(this.btnBaseDataExport2Excel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(172, 52);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 22;
            this.label2.Text = "结束时间：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(172, 17);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 23;
            this.label1.Text = "起始时间：";
            // 
            // dtpStopTime
            // 
            this.dtpStopTime.Location = new System.Drawing.Point(242, 48);
            this.dtpStopTime.Name = "dtpStopTime";
            this.dtpStopTime.Size = new System.Drawing.Size(122, 21);
            this.dtpStopTime.TabIndex = 20;
            // 
            // dtpStartTime
            // 
            this.dtpStartTime.Location = new System.Drawing.Point(242, 12);
            this.dtpStartTime.Name = "dtpStartTime";
            this.dtpStartTime.Size = new System.Drawing.Size(122, 21);
            this.dtpStartTime.TabIndex = 21;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(956, 610);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Main";
            this.Text = "Form1";
            this.groupBox2.ResumeLayout(false);
            this.tabConServerLog.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tpgRegion.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvShowRegion)).EndInit();
            this.tpgDepart.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvShowDepart)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TabControl tabConServerLog;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.RichTextBox ServerLog;
        private System.Windows.Forms.TabPage tpgRegion;
        public System.Windows.Forms.DataGridView dgvShowRegion;
        private System.Windows.Forms.TabPage tpgDepart;
        private System.Windows.Forms.DataGridView dgvShowDepart;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnBaseDataExport2Excel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpStopTime;
        private System.Windows.Forms.DateTimePicker dtpStartTime;
    }
}

