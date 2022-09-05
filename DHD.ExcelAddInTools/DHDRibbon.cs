using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:   按照以下步骤启用功能区(XML)项:

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new DHDRibbon();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace DHD.ExcelAddInTools
{
    [ComVisible(true)]
    public class DHDRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DHDRibbon()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("DHD.ExcelAddInTools.DHDRibbon.xml");
        }

        #endregion

        #region 功能区回调
        //在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion


        #region 公共处理
        /// <summary>
        /// 获取控件图标回调
        /// </summary>
        /// <param name="ctrl"></param>
        /// <returns></returns>
        public System.Drawing.Bitmap Get_ControlImage(Office.IRibbonControl ctrl)
        {
            switch (ctrl.Id)
            {
                // 插入日期时间部分按钮
                case "DHD_BTN_InsertDate":
                case "DHD_ContextMenuCell_InsertDate":
                case "DHD_ContextMenuListRange_InsertDate":
                    return new System.Drawing.Bitmap(Properties.Resources.Today_32x32);
                case "DHD_BTN_InsertTime":
                case "DHD_ContextMenuCell_InsertTime":
                case "DHD_ContextMenuListRange_InsertTime":
                    return new System.Drawing.Bitmap(Properties.Resources.Time_32x32);
                case "DHD_BTN_InsertDateTime":
                    return new System.Drawing.Bitmap(Properties.Resources.Calendar_32x32);
                case "DHD_BTN_Calendar":
                    return new System.Drawing.Bitmap(Properties.Resources.SwitchTimeScalesTo_32x32);

                // 显示设置
                case "DHD_Toggle_ShowMainPan":
                    return new System.Drawing.Bitmap(Properties.Resources.Show_32x32);

                // 工作表工作簿操作部分按钮
                case "DHD_BTN_ExportSheetsToFile":  // 导出工作表为单文件
                    return new System.Drawing.Bitmap(Properties.Resources.Export_32x32);
                case "DHD_BTN_MergeSheets":
                    return new System.Drawing.Bitmap(Properties.Resources.AddNewDataSource_32x32);
                case "DHD_BTN_SortSheet":
                    return new System.Drawing.Bitmap(Properties.Resources.SortAsc_32x32);
                case "DHD_BTN_Start_Calculator":
                    return Properties.Resources.Calculator.ToBitmap();

                default:
                    return new System.Drawing.Bitmap(Properties.Resources.settings_32);
            }
        }


        /// <summary>
        /// 向选中的单元格批量写入具体的值
        /// </summary>
        /// <param name="val"></param>
        public void WriteCells(object val)
        {
            if (val == null) return;
            Microsoft.Office.Interop.Excel.Range selRang = Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;

            Int32 cellTotal = 0;
            if (selRang != null && selRang.Cells.Count > 0)
            {
                cellTotal = selRang.Cells.Count;
                for (Int32 i = 1; i <= cellTotal; i++)
                {
                    // 主要任务处理
                    Microsoft.Office.Interop.Excel.Range c = (Microsoft.Office.Interop.Excel.Range)selRang.Cells[i];
                    c.Value = val;
                }
            }

        }

        #endregion


        #region 文本处理

        public void OnClick_Text(Office.IRibbonControl ctrl)
        {
            try
            {
                switch (ctrl.Id)
                {
                    case "DHD_BTN_MergeCellText":
                        Controls.frmMergeCellText frm = new Controls.frmMergeCellText();
                        frm.Show();
                        break;
                    case "DHD_BTN_SeparateCellText":
                        Controls.frmSeparateCellText frm1 = new Controls.frmSeparateCellText();
                        frm1.Show();
                        break;
                    default:
                        MsgBox.Show("没有定义的处理分支！");
                        break;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show($"文本处理启动发生异常：{ex.Message},{ex.StackTrace}", MsgBox.MsgType.Error);
            }
        }


        #endregion

        #region 插入内容

        public void OnClick_InsertDateTime(Office.IRibbonControl ctrl)
        {
            switch (ctrl.Id)
            {
                case "DHD_BTN_InsertDate":
                case "DHD_ContextMenuCell_InsertDate":
                case "DHD_ContextMenuListRange_InsertDate":
                    WriteCells(DateTime.Now.ToString("yyyy-MM-dd"));
                    //Globals.ThisAddIn.Application.OnUndo("撤销 插入GUID", "UndoEE");
                    break;
                case "DHD_BTN_InsertTime":
                case "DHD_ContextMenuCell_InsertTime":
                case "DHD_ContextMenuListRange_InsertTime":
                    WriteCells(DateTime.Now.ToString("HH:mm:ss"));
                    break;
                case "DHD_BTN_InsertDateTime":
                    WriteCells(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    break;
                case "DHD_BTN_Calendar":
                    //Controls.DateTimePicker dtp = new Controls.DateTimePicker();
                    //if (dtp.ShowDialog() == DialogResult.OK)
                    //{
                    //    WriteCells(dtp.Date);
                    //}

                    //Globals.ThisAddIn.Application.InputBox("", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing);
                    break;
                default:
                    break;
            }
        }


        #endregion


        #region 开发调试

        public void OnClick_DEV(Office.IRibbonControl ctrl)
        {
            switch (ctrl.Id)
            {
                case "DHD_DEV_ShowAssemblyInfo":
                    StringBuilder sb = new StringBuilder();
                    Assembly assembly = Assembly.GetExecutingAssembly();
                    sb.AppendLine($"{nameof(assembly.Location)}\t{assembly.Location}");

                    MsgBox.Show(sb.ToString());
                    break;
                default:
                    break;
            }
        }

        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
