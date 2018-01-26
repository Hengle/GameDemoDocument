using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSFrameWork;
using CSVFrameWork;
using XmlFrameWork;
using ReadExcel;
using System.Windows.Forms;
using ExcelToCSV_XML.Path;

namespace ExcelToCSV_XML.Export
{
    class Export
    {

        public void ReadWrite(string readPath, string savePath)
        {
            if (!IsAvalidPath(readPath, savePath))
            {
                return;
            }

            ReadExcelClass readExcelClass = new ReadExcelClass();

            List<DataTableClass> dataTableClassList = readExcelClass.Read(readPath);

            { // 导出 CSV
                Dictionary<int, List<List<string>>> dataDic = AnalysisCSVClass.AnalysisExcel(dataTableClassList);
                WriteCSVClass writeCSVClass = new WriteCSVClass();
                writeCSVClass.CreateCSV(dataDic, savePath);
            }

            { // 导出 XML
                Dictionary<int, List<List<string>>> dataDic = AnalysisXMLClass.AnalysisExcel(dataTableClassList);
                WriteXmlClass writeXmlClass = new WriteXmlClass();
                writeXmlClass.CreateXML(dataDic, savePath);
            }

            { // 导出 CS
                Dictionary<string, StringBuilder> fileStringBuilderDic = AnalysisCS.Analysis(dataTableClassList);
                WriteCSClass writeCSClass = new WriteCSClass();
                writeCSClass.SaveCsToFile(fileStringBuilderDic, savePath);
            }

            FilePathController.ReplacePathFile(readPath, savePath);
        }

        private bool IsAvalidPath(string filePath, string savePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("请选择需要导出的文件");
                return false;
            }

            if (string.IsNullOrEmpty(savePath))
            {
                MessageBox.Show("请选择保存的路径");
                return false;
            }

            return true;
        }
    }
}
