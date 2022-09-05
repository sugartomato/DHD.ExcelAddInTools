namespace DHD.ExcelAddInTools.Controls
{
    partial class frmMergeCellText
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSeparator = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.chkIgnoreBlankCell = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.rsTarget = new DHD.ExcelAddInTools.Controls.RangeSelector();
            this.rsSource = new DHD.ExcelAddInTools.Controls.RangeSelector();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "合并区域：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "存放位置：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 3;
            this.label3.Text = "分隔符：";
            // 
            // txtSeparator
            // 
            this.txtSeparator.Location = new System.Drawing.Point(83, 62);
            this.txtSeparator.Name = "txtSeparator";
            this.txtSeparator.Size = new System.Drawing.Size(51, 21);
            this.txtSeparator.TabIndex = 2;
            this.txtSeparator.Text = ",";
            // 
            // btnOK
            // 
            this.btnOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.btnOK.Location = new System.Drawing.Point(102, 108);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "拼合(&M)";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // chkIgnoreBlankCell
            // 
            this.chkIgnoreBlankCell.AutoSize = true;
            this.chkIgnoreBlankCell.Checked = true;
            this.chkIgnoreBlankCell.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreBlankCell.Location = new System.Drawing.Point(140, 65);
            this.chkIgnoreBlankCell.Name = "chkIgnoreBlankCell";
            this.chkIgnoreBlankCell.Size = new System.Drawing.Size(96, 16);
            this.chkIgnoreBlankCell.TabIndex = 3;
            this.chkIgnoreBlankCell.Text = "忽略空单元格";
            this.chkIgnoreBlankCell.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.button1.Location = new System.Drawing.Point(183, 108);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 30);
            this.button1.TabIndex = 5;
            this.button1.Text = "关闭(&C)";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // rsTarget
            // 
            this.rsTarget.Address = "";
            this.rsTarget.Location = new System.Drawing.Point(83, 35);
            this.rsTarget.Name = "rsTarget";
            this.rsTarget.SheetName = "";
            this.rsTarget.Size = new System.Drawing.Size(150, 21);
            this.rsTarget.TabIndex = 1;
            // 
            // rsSource
            // 
            this.rsSource.Address = "";
            this.rsSource.Location = new System.Drawing.Point(83, 8);
            this.rsSource.Name = "rsSource";
            this.rsSource.SheetName = "";
            this.rsSource.Size = new System.Drawing.Size(150, 21);
            this.rsSource.TabIndex = 0;
            // 
            // frmMergeCellText
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(260, 143);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.chkIgnoreBlankCell);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtSeparator);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.rsTarget);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rsSource);
            this.Controls.Add(this.label3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMergeCellText";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "单元格文本拼合";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmMergeCellText_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private RangeSelector rsSource;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private RangeSelector rsTarget;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSeparator;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.CheckBox chkIgnoreBlankCell;
        private System.Windows.Forms.Button button1;
    }
}