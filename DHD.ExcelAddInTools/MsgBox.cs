using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DHD.ExcelAddInTools
{
    internal class MsgBox
    {
        public static void Show( String msg)
        {
            ShowMsg(msg);
        }

        public static void Show(String msg, MsgType msgType)
        {
            ShowMsg(msg, msgType);
        }


        private static void ShowMsg(String msg, MsgType msgType = MsgType.Normal)
        {
            System.Windows.Forms.MessageBox.Show(msg);
        }



        public enum MsgType
        { 
            Normal,
            Success,
            Error,
            Warning,
            Information,
        }


    }
}
