using System;
using System.Collections.Generic;
using ReadExcel;
using System.Text;

namespace CSFrameWork
{
    class AnalysisCS
    {
        private static Dictionary<string, StringBuilder> _classCustom = new Dictionary<string, StringBuilder>();
        public static Dictionary<string, StringBuilder> Analysis(List<DataTableClass> dataTableClassList)
        {
            for (int i = 0; i < dataTableClassList.Count; ++i)
            {
                GetCSAttributeFile(dataTableClassList[i]);
            }

            return _classCustom;
        }

        private static void GetCSAttributeFile(DataTableClass dataTableClass)
        {
            StringBuilder CSContentStr = new StringBuilder();
            List<List<string>> dataList = new List<List<string>>();

            int columns = dataTableClass.Cols;
            int rows = dataTableClass.Rows;

            string _fileName = dataTableClass.GetValue(0, 0).ToString();
            //将文件名存入到第一个数据
            List<string> fileNameList = new List<string>();
            fileNameList.Add(_fileName);
            dataList.Add(fileNameList);

            //第三行第三列存储的是数据类型
            CSContentStr.AppendLine("using System;");
            CSContentStr.AppendLine("using System.Collections.Generic;");
            CSContentStr.AppendLine("using System.Text;");
            string className = ProcessName(_fileName);
            CSContentStr.AppendLine("public class " + className + "{");

            //遍历所有属性名（第二行，第二列以后整行数据）
            List<string> AttributeList = new List<string>();
            for (int i = 1; i < columns; ++i)
            {
                string value = dataTableClass.GetValue( 2, i).ToString();

                AttributeList.Add(value);

                //第三行数据数据结构处理
                //解析前三行 生成表结构
                string paramDesc = dataTableClass.GetValue( 0, i).ToString();
                string paramType = dataTableClass.GetValue( 2, i).ToString();
                paramType = GetAvalidType(paramType);
                string paramName = dataTableClass.GetValue( 1, i).ToString();

                if (string.IsNullOrEmpty(paramName) || paramName.CompareTo("") == 0)
                {
                    continue;
                }

                CSContentStr.AppendLine("\t/// <summary>");
                CSContentStr.AppendLine("\t///" + paramDesc);
                CSContentStr.AppendLine("\t/// <summary>");
                //当前列的结构语句
                CSContentStr.AppendLine(string.Format("\tpublic {0} {1} ", paramType, paramName) + "{get;private set;}");
            }
            CSContentStr.AppendLine("}");
            // 第二个数据保存属性值
            dataList.Add(AttributeList);
            _classCustom.Add(className, CSContentStr);
        }

        private static string ProcessName(string fileName)
        {
            string result = string.Empty;//定义一个空字符串
            if (fileName.Contains("_"))
            {
                string[] strArray = fileName.Split('_');

                foreach (string s in strArray)//循环处理数组里面每一个字符串
                {
                    result += System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(s);
                }
            }
            else
            {
                result = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(fileName);
            }
            return result + "Data";
        }

        /// <summary>
        /// 获取有效参数类型：只支持 int、long、float、double、string
        /// </summary>
        /// <param name="paramType"></param>
        /// <returns></returns>
        private static readonly Dictionary<string, bool> m_paramTypeDic = new Dictionary<string, bool>() {
            { "int", true},// 、、、long、string
            { "long", true},
            { "float", true},
            { "double", true},
            { "string", true }
        };
        private static string GetAvalidType(string paramType)
        {
            if (m_paramTypeDic.ContainsKey(paramType))
            {
                return paramType;
            }

            return "string";
        }

    }
}
