using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

namespace XmlFrameWork
{
    class WriteXmlClass
    {
        public void CreateXML(Dictionary<int, List<List<string>>> dataDic, string savePath)
        {
            foreach(KeyValuePair<int, List<List<string>>> kv in dataDic)
            {
                CreateTableDataXML(kv.Value, savePath);
            }
        }

        private void CreateTableDataXML(List<List<string>> dataList, string savePath)
        {
            if (dataList == null || dataList.Count <= 2)
            {
                return;
            }

            List<string> fileNameList = dataList[0];
            // 根据类型获取 XML 名
            string xmlName = fileNameList[0];
            string className = "Data";

            // 新建 XML 实例
            XmlDocument xmlDoc = new XmlDocument();

            // 创建跟节点，最上层节点
            XmlElement rootElement = xmlDoc.CreateElement("Root");
            xmlDoc.AppendChild(rootElement);

            List<string> AttributeList = dataList[1];

            for (int i = 2; i < dataList.Count; ++i)
            {
                List<string> rowList = dataList[i];
                XmlElement rowElement = xmlDoc.CreateElement(className);

                for (int j = 0; j < rowList.Count; ++j)
                {
                    string name = AttributeList[j];

                    if (string.IsNullOrEmpty(name))  // 空表列检测
                    {
                        continue;
                    }

                    string value = rowList[j];
                    if (string.IsNullOrEmpty(value))
                    {
                        value = "";
                    }
                    rowElement.SetAttribute(name, value);
                }

                rootElement.AppendChild(rowElement);
            }

            string path = string.Format("{0}\\XML\\{1}.xml", savePath, xmlName);

            if (!Directory.Exists(Path.GetDirectoryName(path)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
            }

            // 已经存在该文件则删除
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            //将XML 文件保存到本地
            xmlDoc.Save(path);
        }
    }
}
