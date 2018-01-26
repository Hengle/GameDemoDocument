using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CSFrameWork
{
    class WriteCSClass
    {
        public void SaveCsToFile(Dictionary<string, StringBuilder> fileStringBuilderDic, string outPath)
        {
            if (!outPath.EndsWith("CS"))
            {
                outPath = string.Format("{0}\\{1}", outPath, "CS");
            }
            SaveToFile(outPath, fileStringBuilderDic, ".cs");
        }

        private void SaveToFile(string outPath, Dictionary<string, StringBuilder> fileStringBuilderDic, string extension)
        {
            if (!Directory.Exists(outPath))
            {
                Directory.CreateDirectory(outPath);
            }
            foreach (var classFile in fileStringBuilderDic)
            {
                string filePath = Path.Combine(outPath, classFile.Key + extension);
                using (FileStream fs = File.Create(filePath))
                {
                    try
                    {
                        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(classFile.Value.ToString());
                        fs.Write(bytes, 0, bytes.Length);
                        fs.Close();
                    }
                    catch (IOException e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
        }
    }
}
