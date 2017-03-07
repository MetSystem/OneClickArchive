using Met.Report;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Met.Report
{
    public class Demo
    {
        private ReportInfo report;
        public Demo()
        {
            report = new ReportInfo();
            report.Title = "NET技术文档";
            report.Author = "met";
            report.Creator = "met";
            report.Keywords = "met,asp.net";
            report.Password = "met";
            report.WaterMarkName = "met";
            report.CoverPath = "PDF/21 分钟 MySQL 入门教程 - v1.0.pdf";
            report.PeportFiles = new List<ReportFileInfo>() {
                new ReportFileInfo() {
                    Name="入门教程",
                    Path = "PDF/21 分钟 MySQL 入门教程 - v1.0.pdf",
                    SubFiles =new List<ReportFileInfo>() {
                        new ReportFileInfo() {
                            Name="客户指南",
                            Path ="PDF/Android Gradle 用户指南 - v1.0.pdf",
                            SubFiles =new List<ReportFileInfo>() {
                                new ReportFileInfo() {
                                    Name="用户指南",
                                    Path = "PDF/Android Gradle 用户指南 - v1.0.pdf"
                                }
                            }
                        },
                    }
                },
                new ReportFileInfo() {Name="OneClickArchive",Path = "PDF/Google Objective-C 风格指南 - v1.0.pdf"}
            };
        }

        /// <summary>
        /// 报告生成水印文字
        /// </summary>
        public void Report1()
        {
            var reportManager = ReportManager.Instace;
            // 生成报告
            reportManager.BuildReport(report, "newpdf.pdf");

            if (!string.IsNullOrEmpty(report.WaterMarkName))
            {
                if (string.IsNullOrEmpty(report.Password))
                {
                    reportManager.SetWatermarkByWord("newpdf.pdf", "newpdf1.pdf", report.WaterMarkName, true);
                }
                else
                {
                    reportManager.SetWatermarkByWord("newpdf.pdf", "newpdf1.pdf", report.WaterMarkName, report.Password, report.Password, 0, true);
                }
            }
        }

        /// <summary>
        /// 报告生成水印图片
        /// </summary>
        public void Report2()
        {
            var reportManager = ReportManager.Instace;
            // 生成报告
            reportManager.BuildReport(report, "newpdf.pdf");

            if (!string.IsNullOrEmpty(report.WaterMarkName))
            {
                if (string.IsNullOrEmpty(report.Password))
                {
                    reportManager.SetWatermarkByImage("newpdf.pdf", "newpdf1.pdf", report.WaterMarkName, 0, 0);
                }
                else
                {
                    reportManager.SetWatermarkByImage("newpdf.pdf", "newpdf1.pdf", report.WaterMarkName, 0, 0, report.Password, report.Password, 0);
                }
            }
        }
    }
}
