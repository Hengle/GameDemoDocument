using System;
using System.Collections.Generic;
using ReadExcel;

namespace CSVFrameWork
{
    class AnalysisCSVClass
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
            //将文件名存入到第一个数据
            string _fileName = dataTableClass.GetValue(0, 0).ToString();

            List<string> fileNameList = new List<string>();
            fileNameList.Add(_fileName);
            dataList.Add(fileNameList);

            int rows = dataTableClass.Rows;
            int columns = dataTableClass.Cols;
            // 遍历前三行
            // 第一行 属性说明
            // 第二行 属性名
            // 第三行 类型声明(int, string、、)
            for (int i = 0; i < rows; ++i)
            {
                List<string> AttributeList = new List<string>();
                bool isAvalidRow = true;
                for (int j = 1; j < columns; ++j)
                {
                    string value = dataTableClass.GetValue(i, j).ToString();

                    if (j == 1 && i >= 3 && string.IsNullOrEmpty(value))
                    {
                        isAvalidRow = false;
                        break;
                    }
                    AttributeList.Add(value);
                }
                if (isAvalidRow)
                {
                    dataList.Add(AttributeList);
                }
            }

            dataDic[index] = dataList;

            return dataDic;
        }
    }
}
