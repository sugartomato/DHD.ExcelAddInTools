using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace DHD.ExcelAddInTools
{
    internal class Config
    {
        public static Dictionary<string, Microsoft.Office.Tools.CustomTaskPane> TaskPans = new Dictionary<string, Microsoft.Office.Tools.CustomTaskPane>();
        public static Microsoft.Office.Tools.CustomTaskPane GetCustomTaskPane(String panID, string panTitle, System.Windows.Forms.UserControl ctrl)
        {
            string key = string.Format("{0}({1})", panID, Globals.ThisAddIn.Application.Hwnd);
            if (!TaskPans.ContainsKey(key))
            {
                var pan = Globals.ThisAddIn.CustomTaskPanes.Add(ctrl, panTitle);
                TaskPans[key] = pan;
            }
            return TaskPans[key];
        }

        public static Microsoft.Office.Tools.CustomTaskPane GetActiveWindowMainTaksPane()
        {
            String key = "辅助面板[" + Globals.ThisAddIn.Application.Hwnd.ToString() + "]";
            return TaskPans[key];
        }

        #region 配置属性


        /// <summary>
        /// 固定的Excel文件位置列表
        /// </summary>
        public static List<Model.PinnedFile> PinnedFiles
        {
            get
            {
                List<Model.PinnedFile> result = null;
                String xPath = "/DHD/PinnedFiles";
                XmlNode node = GetNode(xPath);
                if (node == null) return null;
                XmlNodeList list = node.SelectNodes("./*");
                if (list != null && list.Count > 0)
                {
                    result = new List<Model.PinnedFile>();
                    foreach (XmlNode n in list)
                    {
                        Model.PinnedFile tmpFile = new Model.PinnedFile();
                        tmpFile.FileName = n.SelectSingleNode("./FileName").InnerText;
                        tmpFile.FilePath = n.SelectSingleNode("./FilePath").InnerText;
                        tmpFile.Mark = n.SelectSingleNode("./Mark").InnerText;
                        result.Add(tmpFile);
                    }
                }
                return result;
            }
            set
            {
                String xPath = "/DHD/PinnedFiles";
                XmlNode node = GetNode(xPath);
                if (node == null)
                {
                    node = XmlDoc.CreateNode(XmlNodeType.Element, "PinnedFiles", "");
                    XmlDoc.DocumentElement.AppendChild(node);
                }
                node.RemoveAll();

                if (value != null || value.Count > 0)
                {
                    for (Int32 i = 0; i < value.Count; i++)
                    {
                        XmlNode nodeFile = XmlDoc.CreateNode(XmlNodeType.Element, "File", "");
                        node.AppendChild(nodeFile);

                        XmlNode nodeFileName = XmlDoc.CreateNode(XmlNodeType.Element, "FileName", "");
                        nodeFileName.InnerText = value[i].FileName;
                        nodeFile.AppendChild(nodeFileName);

                        XmlNode nodeFilePath = XmlDoc.CreateNode(XmlNodeType.Element, "FilePath", "");
                        nodeFilePath.InnerText = value[i].FilePath;
                        nodeFile.AppendChild(nodeFilePath);

                        XmlNode nodeFileMark = XmlDoc.CreateNode(XmlNodeType.Element, "Mark", "");
                        nodeFileMark.InnerText = value[i].Mark;
                        nodeFile.AppendChild(nodeFileMark);
                    }
                }
            }
        }

        /// <summary>
        /// 上次选择的目录的路径。如果该路径不存在，则返回桌面
        /// </summary>
        public static String LastSelectDir
        {
            get
            {
                String defaultPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                String xPath = "/DHD/LastSelectDir";
                XmlNode node = GetNode(xPath);
                if (node == null) return defaultPath;
                String result = node.Value;
                if (!System.IO.Directory.Exists(result)) return result;
                return defaultPath;

            }
            set
            {
                String xPath = "/DHD/LastSelectDir";
                XmlNode node = GetNode(xPath);
                if (node == null) return;
                node.Value = value;
            }
        }

        #endregion

        #region XML文件操作

        private static String ConfigFilePath
        {
            get
            {
                String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DHD\DHDExcelAddInTools\";
                if (path.StartsWith("file:"))
                {
                    path = System.IO.Path.GetFullPath(path.Replace("file:///", ""));
                }
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }

                path = path + @"\Config.xml";

                return path.Replace(@"\\", "\\");
            }
        }

        /// <summary>
        /// 获取配置文件路径
        /// </summary>
        public static String FilePath
        {
            get
            {
                return ConfigFilePath;
            }
        }
        private static XmlNode GetNode(String path)
        {
            XmlNode node = XmlDoc.SelectSingleNode(path);
            return node;

        }

        private static String GetNodeValue(String path)
        {
            XmlNode node = XmlDoc.SelectSingleNode(path);
            return node?.InnerText?.Trim();
        }

        private static void SetNodeValue(String path, String value)
        {
            XmlNode node = XmlDoc.SelectSingleNode(path);
            if (node == null)
            {

            }
            if (node != null) node.InnerText = value;
        }

        private static XmlDocument _XmlDoc = null;
        /// <summary>
        /// XML配置文件。
        /// 只加载一次到内存中。后续如果有修改，重新加载
        /// </summary>
        private static XmlDocument XmlDoc
        {
            get
            {
                if (_XmlDoc == null)
                    LoadCoinfig();
                return _XmlDoc;
            }
        }

        /// <summary>
        /// 加载配置
        /// </summary>
        public static void LoadCoinfig()
        {
            _XmlDoc = new XmlDocument();
            // 配置文件不存在，则创建配置文件
            if (!System.IO.File.Exists(ConfigFilePath))
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
                sb.AppendLine("<DHD>");
                sb.AppendLine("<ShowSpotlight>false</ShowSpotlight>");
                sb.AppendLine("</DHD>");

                _XmlDoc.LoadXml(sb.ToString());
                _XmlDoc.Save(ConfigFilePath);
            }
            else
            {
                try
                {
                    _XmlDoc.Load(ConfigFilePath);
                }
                catch (Exception ex)
                {
                    throw new ApplicationException("加载配置文件异常：" + ex.Message);
                }
            }
        }

        /// <summary>
        /// 保存配置到配置文件
        /// </summary>
        public static void SaveConfig()
        {
            XmlDoc.Save(ConfigFilePath);
        }

        #endregion

    }
}
