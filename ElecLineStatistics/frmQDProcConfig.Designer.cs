namespace ElecStatistics
{
    partial class frmQDProcConfig
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnUpdateHeader = new System.Windows.Forms.Button();
            this.cmbSAPHeader = new System.Windows.Forms.ComboBox();
            this.lvwProcConfig = new System.Windows.Forms.ListView();
            this.colSAPHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colRepHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnSaveSequence = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightBlue;
            this.panel1.Controls.Add(this.btnSaveSequence);
            this.panel1.Controls.Add(this.btnUpdateHeader);
            this.panel1.Controls.Add(this.cmbSAPHeader);
            this.panel1.Controls.Add(this.lvwProcConfig);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(7);
            this.panel1.Size = new System.Drawing.Size(297, 400);
            this.panel1.TabIndex = 0;
            // 
            // btnUpdateHeader
            // 
            this.btnUpdateHeader.Location = new System.Drawing.Point(7, 358);
            this.btnUpdateHeader.Name = "btnUpdateHeader";
            this.btnUpdateHeader.Size = new System.Drawing.Size(132, 32);
            this.btnUpdateHeader.TabIndex = 2;
            this.btnUpdateHeader.Text = "更新表头";
            this.btnUpdateHeader.UseVisualStyleBackColor = true;
            // 
            // cmbSAPHeader
            // 
            this.cmbSAPHeader.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSAPHeader.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmbSAPHeader.FormattingEnabled = true;
            this.cmbSAPHeader.Location = new System.Drawing.Point(7, 320);
            this.cmbSAPHeader.Name = "cmbSAPHeader";
            this.cmbSAPHeader.Size = new System.Drawing.Size(283, 32);
            this.cmbSAPHeader.TabIndex = 1;
            // 
            // lvwProcConfig
            // 
            this.lvwProcConfig.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colRepHeader,
            this.colSAPHeader});
            this.lvwProcConfig.Dock = System.Windows.Forms.DockStyle.Top;
            this.lvwProcConfig.FullRowSelect = true;
            this.lvwProcConfig.GridLines = true;
            this.lvwProcConfig.HideSelection = false;
            this.lvwProcConfig.Location = new System.Drawing.Point(7, 7);
            this.lvwProcConfig.Name = "lvwProcConfig";
            this.lvwProcConfig.Size = new System.Drawing.Size(283, 307);
            this.lvwProcConfig.TabIndex = 0;
            this.lvwProcConfig.UseCompatibleStateImageBehavior = false;
            this.lvwProcConfig.View = System.Windows.Forms.View.Details;
            // 
            // colSAPHeader
            // 
            this.colSAPHeader.Text = "原始列头";
            this.colSAPHeader.Width = 131;
            // 
            // colRepHeader
            // 
            this.colRepHeader.Text = "报表列头";
            this.colRepHeader.Width = 122;
            // 
            // btnSaveSequence
            // 
            this.btnSaveSequence.Location = new System.Drawing.Point(160, 358);
            this.btnSaveSequence.Name = "btnSaveSequence";
            this.btnSaveSequence.Size = new System.Drawing.Size(130, 32);
            this.btnSaveSequence.TabIndex = 3;
            this.btnSaveSequence.Text = "保存设置";
            this.btnSaveSequence.UseVisualStyleBackColor = true;
            // 
            // frmQDProcConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(297, 400);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmQDProcConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "配置转换流程";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ColumnHeader colSAPHeader;
        private System.Windows.Forms.ColumnHeader colRepHeader;
        public System.Windows.Forms.ListView lvwProcConfig;
        public System.Windows.Forms.ComboBox cmbSAPHeader;
        public System.Windows.Forms.Button btnUpdateHeader;
        public System.Windows.Forms.Button btnSaveSequence;
    }
}