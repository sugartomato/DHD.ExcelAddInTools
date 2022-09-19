using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DHD.ExcelAddInTools.Model
{
    internal class PinnedFile
    {
        public PinnedFile() { }
        public PinnedFile(String filePath)
        {
            this.FilePath = filePath;
            try
            {
                this.FileName = System.IO.Path.GetFileName(filePath);
            }
            catch (Exception)
            {
                this.FileName = "未能通过路径获取到文件名！";
            }
            this.Mark = System.Guid.NewGuid().ToString("N").ToUpper();
        }


        public String FileName { get; set; }
        public String FilePath { get; set; }
        public String Mark { get; set; }
    }
}
