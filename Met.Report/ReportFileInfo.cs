using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Met.Report
{
    /// <summary>
    /// 报告文件
    /// </summary>
    public class ReportFileInfo
    {
        private string _name;

        /// <summary>
        /// 名称
        /// </summary>
        public string Name
        {
            get
            {
                if (_name == null)
                {
                    FileInfo info = new FileInfo(Path);
                    _name = info.Name.Replace(info.Extension, string.Empty);
                }
                return _name;
            }
            set
            {
                this._name = value;
            }
        }

        /// <summary>
        /// 路径
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// 报告文件
        /// </summary>
        public List<ReportFileInfo> SubFiles { get; set; }
    }

}
