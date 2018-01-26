using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelToCSV_XML.Path;

using ExcelToCSV_XML.Export;

namespace ReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            pathList = FilePathController.PathList();
            AddEvent();
        }

        public Form1(string[] args)
        {
            if (args.Length <= 1)
            {
                return;
            }

            string savePath = args[0];
            for (int i = 1; i < args.Length; ++i)
            {
                Export export = new Export();
                export.ReadWrite(args[i], savePath);
            }

            Environment.Exit(0);
        }

        private void AddEvent()
        {
            this.PathText.DragEnter += PathText_DragEnter;
            this.PathText.DragDrop += PathText_DragDrop;

            this.SavePath.DragEnter += SavePath_DragEnter;
            this.SavePath.DragDrop += SavePath_DragDrop;
        }

        private void Save_Click(object sender, EventArgs e)
        {
            Export export = new Export();
            export.ReadWrite(this.PathText.Text, this.SavePath.Text);
        }

        private void PathText_TextChanged(object sender, EventArgs e) { }

        private void PathText_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void PathText_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            this.PathText.Text = path;
        }

        private List<string> pathList = new List<string>();
        private void SelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            
            if (pathList.Count >= 1)
            {
                openFileDialog.InitialDirectory = Path.GetDirectoryName(pathList[0]);
            }
            else
            {
                openFileDialog.InitialDirectory = "d:\\";//注意这里写路径时要用c:\\而不是c:\
            }

            openFileDialog.Filter = "Excel文件|*.xlsx*"; // "Excel文件|*.xlsx*|Excel文件|*.xls*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.PathText.Text = openFileDialog.FileName;
            }
        }

        private void SelectSavePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择保存XML路径";
            if (pathList.Count >= 2)
            {
                dialog.SelectedPath = pathList[1];
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                if (string.IsNullOrEmpty(foldPath))
                {
                    MessageBox.Show("请选择有效的路径");
                    return;
                }

                this.SavePath.Text = foldPath;
            }
        }

        private void SavePath_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void SavePath_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            this.SavePath.Text = Path.GetDirectoryName(path);
        }
    }
}