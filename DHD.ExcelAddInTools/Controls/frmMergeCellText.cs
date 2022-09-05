using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace DHD.ExcelAddInTools.Controls
{
    public partial class frmMergeCellText : Form
    {
        public frmMergeCellText()
        {
            InitializeComponent();

            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.ShowIcon = false;
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void frmMergeCellText_Load(object sender, EventArgs e)
        {
            Range rng = Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
            rsSource.Address = rng.Address;
            rsSource.SheetName = rng.Worksheet.Name;

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (String.IsNullOrEmpty(rsTarget.Address) || rsTarget.Address.Length == 0)
                {
                    MessageBox.Show("请设置存储位置！");
                    return;
                }

                String result = String.Empty;


                Worksheet sheet = ((Worksheet)Globals.ThisAddIn.Application.Sheets[rsSource.SheetName]);

                // 拆分地址
                String[] address = rsSource.Address.Split(',');
                foreach (String add in address)
                {
                    Range rng = sheet.Range[add];
                    foreach (Range cell in rng.Cells)
                    {
                        String val = Convert.ToString(cell.Value);
                        if (String.IsNullOrEmpty(val) && !chkIgnoreBlankCell.Checked == true)
                        {
                            continue;
                        }
                        else
                        {
                            result += val + txtSeparator.Text;
                        }
                    }
                }

                result = result.Substring(0, result.Length - txtSeparator.Text.Length);
                ((Worksheet)Globals.ThisAddIn.Application.Sheets[rsTarget.SheetName]).Range[rsTarget.Address].Value = result;
                MsgBox.Show("操作完成！", MsgBox.MsgType.Success);
            }
            catch (Exception ex)
            {
                MsgBox.Show($"操作异常：{ex.Message}", MsgBox.MsgType.Error);
                //Common.ShowError("合并异常：" + ex.Message);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
