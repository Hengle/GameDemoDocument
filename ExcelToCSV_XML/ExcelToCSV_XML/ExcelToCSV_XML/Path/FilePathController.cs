using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReadTXT;

namespace ExcelToCSV_XML.Path
{
    class FilePathController
    {
        private static readonly string pathFile = "Path.txt";
        private static List<string> pathList = new List<string>();
        public static List<string> PathList()
        {
            string path = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            ReadOpenPathFile readOpenPathFile = new ReadOpenPathFile();
            pathList.Clear();
            pathList = readOpenPathFile.Read(path + pathFile);
            return pathList;
        }

        public static void ReplacePathFile(string filePath, string savePath)
        {
            ReadOpenPathFile readOpenPathFile = new ReadOpenPathFile();
            string path = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            path = path + pathFile;
            readOpenPathFile.Delete(path);

            List<string> pathList = new List<string>();
            pathList.Add(filePath);
            pathList.Add(savePath);

            readOpenPathFile.Write(path, pathList);
        }
    }
}
