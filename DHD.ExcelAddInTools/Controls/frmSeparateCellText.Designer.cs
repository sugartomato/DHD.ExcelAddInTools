namespace DHD.ExcelAddInTools.Controls
{
    partial class frmSeparateCellText
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
            this.btnClose = new System.Windows.Forms.Button();
            this.chkIgnoreBlankCell = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtSeparator = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.rsTarget = new DHD.ExcelAddInTools.Controls.RangeSelector();
            this.label1 = new System.Windows.Forms.Label();
            this.rsSource = new DHD.ExcelAddInTools.Controls.RangeSelector();
            this.label3 = new System.Windows.Forms.Label();
            this.rbtHorizontal = new System.Windows.Forms.RadioButton();
            this.rbtVertical = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnClose.Location = new System.Drawing.Point(178, 124);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 30);
            this.btnClose.TabIndex = 14;
            this.btnClose.Text = "关闭(&C)";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // chkIgnoreBlankCell
            // 
            this.chkIgnoreBlankCell.AutoSize = true;
            this.chkIgnoreBlankCell.Checked = true;
            this.chkIgnoreBlankCell.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreBlankCell.Location = new System.Drawing.Point(152, 62);
            this.chkIgnoreBlankCell.Name = "chkIgnoreBlankCell";
            this.chkIgnoreBlankCell.Size = new System.Drawing.Size(96, 16);
            this.chkIgnoreBlankCell.TabIndex = 10;
            this.chkIgnoreBlankCell.Text = "忽略空单元格";
            this.chkIgnoreBlankCell.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.btnOK.Location = new System.Drawing.Point(97, 124);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 13;
            this.btnOK.Text = "拆分(&S)";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtSeparator
            // 
            this.txtSeparator.Location = new System.Drawing.Point(95, 59);
            this.txtSeparator.Name = "txtSeparator";
            this.txtSeparator.Size = new System.Drawing.Size(51, 21);
            this.txtSeparator.TabIndex = 9;
            this.txtSeparator.Text = ",";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 11;
            this.label2.Text = "存放位置：";
            // 
            // rsTarget
            // 
            this.rsTarget.Address = "";
            this.rsTarget.Location = new System.Drawing.Point(95, 32);
            this.rsTarget.Name = "rsTarget";
            this.rsTarget.SheetName = "";
            this.rsTarget.Size = new System.Drawing.Size(150, 21);
            this.rsTarget.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "拆分单元格：";
            // 
            // rsSource
            // 
            this.rsSource.Address = "";
            this.rsSource.Location = new System.Drawing.Point(95, 5);
            this.rsSource.Name = "rsSource";
            this.rsSource.SheetName = "";
            this.rsSource.Size = new System.Drawing.Size(150, 21);
            this.rsSource.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "分隔符：";
            // 
            // rbtHorizontal
            // 
            this.rbtHorizontal.AutoSize = true;
            this.rbtHorizontal.Location = new System.Drawing.Point(18, 93);
            this.rbtHorizontal.Name = "rbtHorizontal";
            this.rbtHorizontal.Size = new System.Drawing.Size(71, 16);
            this.rbtHorizontal.TabIndex = 15;
            this.rbtHorizontal.Text = "水平填充";
            this.rbtHorizontal.UseVisualStyleBackColor = true;
            // 
            // rbtVertical
            // 
            this.rbtVertical.AutoSize = true;
            this.rbtVertical.Checked = true;
            this.rbtVertical.Location = new System.Drawing.Point(97, 93);
            this.rbtVertical.Name = "rbtVertical";
            this.rbtVertical.Size = new System.Drawing.Size(71, 16);
            this.rbtVertical.TabIndex = 16;
            this.rbtVertical.TabStop = true;
            this.rbtVertical.Text = "垂直填充";
            this.rbtVertical.UseVisualStyleBackColor = true;
            // 
            // frmSeparateCellText
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(265, 166);
            this.Controls.Add(this.rbtVertical);
            this.Controls.Add(this.rbtHorizontal);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.chkIgnoreBlankCell);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtSeparator);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.rsTarget);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rsSource);
            this.Controls.Add(this.label3);
            this.Name = "frmSeparateCellText";
            this.Text = "frmSeparateCellText";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmSeparateCellText_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.CheckBox chkIgnoreBlankCell;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtSeparator;
        private System.Windows.Forms.Label label2;
        private RangeSelector rsTarget;
        private System.Windows.Forms.Label label1;
        private RangeSelector rsSource;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RadioButton rbtHorizontal;
        private System.Windows.Forms.RadioButton rbtVertical;
    }
}