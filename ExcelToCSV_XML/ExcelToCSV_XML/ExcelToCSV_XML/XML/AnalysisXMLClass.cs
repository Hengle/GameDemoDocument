using System;
using System.Collections.Generic;
using ReadExcel;

namespace XmlFrameWork
{
    class AnalysisXMLClass
    {
        private static Dictionary<int, List<List<string>>> dataDic = new Dictionary<int, List<List<string>>>();

        public static Dictionary<int, List<List<string>>> AnalysisExcel(List<DataTableClass> dataTableClassList)
        {
            for (int i = 0; i < dataTableClassList.Count; ++i)
            {
                AnalysisExcel(dataTableClassList[i], i);
            }

            return dataDic;
        }

        private static Dictionary<int, List<List<string>>> AnalysisExcel(DataTableClass dataTableClass, int index)
        {
            if (dataTableClass == null)
            {
                return null;
            }

            List<List<string>> dataList = new List<List<string>>();

            //解析需要保存的XML名(第一行第一列的值)
            string _fileName = dataTableClass.GetValue(0, 0).ToString();
            //将文件名存入到第一个数据
            List<string> fileNameList = new List<string>();
            fileNameList.Add(_fileName);
            dataList.Add(fileNameList);

            int columns = dataTableClass.Cols;
            int rows = dataTableClass.Rows;

            //遍历所有属性名（第二行，第二列以后整行数据）
            List<string> AttributeList = new List<string>();
            for (int i = 1; i < columns; ++i)
            {
                string value = dataTableClass.GetValue(1, i).ToString();
                AttributeList.Add(value);
            }
            // 第二个数据保存属性值
            dataList.Add(AttributeList);

            // 表规则为从第四行，第二列开始读数据
            for (int i = 3; i < rows; i++)
            {
                List<string> rowList = new List<string>();
                bool isAvalidRow = true;
                for (int j = 1; j < columns; j++)
                {
                    string value = dataTableClass.GetValue(i, j).ToString();
                    if (j == 1 && i >= 3 && string.IsNullOrEmpty(value))
                    {
                        isAvalidRow = false;
                        break;
                    }
                    rowList.Add(value);
                }
                if (isAvalidRow)
                {
                    dataList.Add(rowList);
                }
            }

            dataDic[index] = dataList;

            return dataDic;
        }
    }
}
