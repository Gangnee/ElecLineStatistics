namespace ElecStatistics
{
    partial class frmLineSelection
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
            this.lvwRTLine = new System.Windows.Forms.ListView();
            this.colLine = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnSelectLine = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightBlue;
            this.panel1.Controls.Add(this.btnSelectLine);
            this.panel1.Controls.Add(this.lvwRTLine);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(260, 378);
            this.panel1.TabIndex = 0;
            // 
            // lvwRTLine
            // 
            this.lvwRTLine.CheckBoxes = true;
            this.lvwRTLine.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colLine,
            this.columnHeader1});
            this.lvwRTLine.GridLines = true;
            this.lvwRTLine.HideSelection = false;
            this.lvwRTLine.Location = new System.Drawing.Point(3, 3);
            this.lvwRTLine.Name = "lvwRTLine";
            this.lvwRTLine.Size = new System.Drawing.Size(253, 324);
            this.lvwRTLine.TabIndex = 0;
            this.lvwRTLine.UseCompatibleStateImageBehavior = false;
            this.lvwRTLine.View = System.Windows.Forms.View.Details;
            // 
            // colLine
            // 
            this.colLine.Text = "线别";
            this.colLine.Width = 74;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "线别备注";
            this.columnHeader1.Width = 158;
            // 
            // btnSelectLine
            // 
            this.btnSelectLine.Location = new System.Drawing.Point(12, 333);
            this.btnSelectLine.Name = "btnSelectLine";
            this.btnSelectLine.Size = new System.Drawing.Size(236, 35);
            this.btnSelectLine.TabIndex = 1;
            this.btnSelectLine.Text = "选择线别";
            this.btnSelectLine.UseVisualStyleBackColor = true;
            // 
            // frmLineSelection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(260, 378);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmLineSelection";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "选择工时线别";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ColumnHeader colLine;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        public System.Windows.Forms.ListView lvwRTLine;
        public System.Windows.Forms.Button btnSelectLine;
    }
}