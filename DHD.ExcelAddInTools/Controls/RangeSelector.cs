using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DHD.ExcelAddInTools.Controls
{
    public partial class RangeSelector : UserControl
    {
        public RangeSelector()
        {
            InitializeComponent();

            // 初始化控件尺寸

            SetControl(150);

        }

        private void RangeSelector_Resize(object sender, EventArgs e)
        {
            if (this.Height > txtAddress.Height)
            {
                this.Height = txtAddress.Height;
            }

            SetControl(this.Width);
        }


        public void SetControl(int width)
        {
            txtAddress.Width = width;
            txtAddress.Location = new Point(0, 0);

            btnSel.Image = DHD.ExcelAddInTools.Properties.Resources.RefEditMin;
            btnSel.Height = txtAddress.Height - 2;
            btnSel.Width = btnSel.Height;
            btnSel.Location = new Point(txtAddress.Width - btnSel.Width - 1, 1);

            this.Width = txtAddress.Width;
            this.Height = txtAddress.Height;
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
            if (this.ParentForm == null) return;

            Form parentFrm = this.ParentForm;
            parentFrm.Visible = false;
            object rng;

            rng = Globals.ThisAddIn.Application.InputBox("选择区域：", "", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

            if (rng != null)
            {

                _propSheetName = (rng as Microsoft.Office.Interop.Excel.Range).Worksheet.Name;
                _propAddress = (rng as Microsoft.Office.Interop.Excel.Range).Address;
                this.txtAddress.Text = _propSheetName + "!" + _propAddress;
            }

            parentFrm.Visible = true;
        }


        private String _propAddress = String.Empty;
        /// <summary>
        /// 所选择区域
        /// </summary>
        public string Address
        {
            get
            {
                return _propAddress;
            }

            set
            {
                if (value != null)
                {
                    _propAddress = value;
                    txtAddress.Text = _propAddress;
                }
            }
        }

        private String _propSheetName = String.Empty;
        /// <summary>
        /// 获取所选择区域所在的工作表名
        /// </summary>
        public String SheetName
        {
            get
            {
                return _propSheetName;
            }
            set
            {
                _propSheetName = value;
            }
        }

    }
}
