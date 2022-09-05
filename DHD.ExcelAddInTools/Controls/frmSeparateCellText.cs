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
    public partial class frmSeparateCellText : Form
    {
        public frmSeparateCellText()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.ShowIcon = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            this.Text = "单元格文本拆分";
        }

        private void frmSeparateCellText_Load(object sender, EventArgs e)
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
                    MessageBox.Show("请设置存储目标单元格！");
                    return;
                }

                List<String> result = new List<string>();
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
                            String[] arr1 = val.Split(new String[] { txtSeparator.Text }, StringSplitOptions.None);
                            if (arr1 != null && arr1.Length > 0)
                            {
                                for (Int32 z = 0; z < arr1.Length; z++)
                                {
                                    result.Add(arr1[z]);
                                }
                            }
                        }
                    }
                }

                // 填充生成的内容
                // 目标地址之后，如何按照顺序赋值
                if (result.Count > 0)
                {
                    Worksheet targetSheet = ((Worksheet)Globals.ThisAddIn.Application.Sheets[rsTarget.SheetName]);
                    Int32 rowIndex = targetSheet.Range[rsTarget.Address].Row;
                    Int32 columnIndex = targetSheet.Range[rsTarget.Address].Column;
                    for (Int32 i = 0; i < result.Count; i++)
                    {
                        ((Range)targetSheet.Cells[rowIndex, columnIndex]).Value = result[i];

                        if (rbtHorizontal.Checked)
                        {
                            columnIndex += 1;
                        }
                        else if (rbtVertical.Checked)
                        {
                            rowIndex += 1;
                        }
                    }
                }
                MsgBox.Show("操作完成！", MsgBox.MsgType.Success);
            }
            catch (Exception ex)
            {
                MsgBox.Show($"操作异常：{ex.Message}", MsgBox.MsgType.Error);
                //Common.WriteConsole("拆分异常：" + ex.Message + ex.StackTrace);
            }

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
