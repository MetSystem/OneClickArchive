using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Met.Report
{
    /// <summary>
    /// 报表信息
    /// </summary>
    public class ReportInfo
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }
        /// <summary>
        /// 创建者
        /// </summary>
        public string Creator { get; set; }
        /// <summary>
        /// 关键字
        /// </summary>
        public string Keywords { get; set; }
        /// <summary>
        /// 主题
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// 文档密码
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// 水印
        /// </summary>
        public string WaterMarkName { get; set; }

        /// <summary>
        /// 封面路径（默认无）
        /// </summary>
        public string CoverPath { get; set; }

        /// <summary>
        /// 页面标题（每一页）
        /// </summary>
        public string PageTitle { get; set; }

        /// <summary>
        ///  文件名称
        /// </summary>
        public List<ReportFileInfo> PeportFiles { get; set; }
    }

}
