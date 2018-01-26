using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.IO.MemoryMappedFiles;
using Excel;
using System.Data.Odbc;

namespace ReadExcel
{
    public class DataTableClass
    {
        private List<List<object>> dataList = new List<List<object>>();

        // 行
        public int Rows {
            get;
            private set;
        }

        // 列
        public int Cols
        {
            get;
            private set;
        }

        public object GetValue(int row, int col)
        {
            if (row >= dataList.Count)
            {
                return null;
            }
            if (col >= dataList[row].Count)
            {
                return null;
            }

            return dataList[row][col];
        }

        public DataTableClass(DataTable dataTable)
        {
            dataList = new List<List<object>>();
            if (dataTable == null)
            {
                Rows = 0;
                Cols = 0;
                return;
            }
            Rows = dataTable.Rows.Count;
            Cols = dataTable.Columns.Count;

            for (int i = 0; i < Rows; ++i)
            {
                List<object> rowList = new List<object>();

                for (int j = 0; j < Cols; ++j)
                {
                    object valueObject = dataTable.Rows[i][j];
                    rowList.Add(valueObject);
                }

                dataList.Add(rowList);
            }
        }
    }

    class ReadExcelClass
    {
        public List<DataTableClass> Read(string path)
        {
            List<DataTable> dataTableList = GetDataTable(path);
            if (dataTableList == null || dataTableList.Count <= 0)
            {
                return null;
            }

            List<DataTableClass> dataTableClassList = new List<DataTableClass>();
            for (int i = 0; i < dataTableList.Count; ++i)
            {
                DataTableClass dataTableClass = new DataTableClass(dataTableList[i]);
                dataTableClassList.Add(dataTableClass);
            }

            return dataTableClassList;
        }

        private List<DataTable> GetDataTable(string path)
        {
            if (Path.GetExtension(path).CompareTo(".xlsx") != 0)
            {
                return null;
            }

            List<DataTable> dataTableList = new List<DataTable>();

            if (Path.GetExtension(path).CompareTo(".xlsx") == 0)
            {
                dataTableList = ReadExcel(path);
            }
            else
            {
                dataTableList = ReadExcelXLS(path);
            }

            return dataTableList;
        }

        private List<DataTable> ReadExcel(string path)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            List<DataTable> dataTableList = new List<DataTable>();
            
            try
            {
                FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                DataSet result = excelReader.AsDataSet();

                if (result.Tables.Count <= 0)
                {
                    return null;
                }

                for (int i = 0; i < result.Tables.Count; ++i)
                {
                    dataTableList.Add(result.Tables[i]);
                }
            }
            catch (Exception e) { throw (e); }

            return dataTableList;
        }

        private List<DataTable> ReadExcelXLS(string path)
        {
            DataTable dtYourData = new DataTable("YourData");

            // Must be saved as excel 2003 workbook, not 2007, mono issue really
            string con = "Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; Dbq=" + path + ";";

            string yourQuery = "SELECT * FROM [Sheet1$]";
            // our odbc connector 
            OdbcConnection oCon = new OdbcConnection(con);
            // our command object 
            OdbcCommand oCmd = new OdbcCommand(yourQuery, oCon);
            // table to hold the data 	
            // open the connection 
            oCon.Open();
            // lets use a datareader to fill that table! 
            OdbcDataReader rData = oCmd.ExecuteReader();
            // now lets blast that into the table by sheer man power! 
            dtYourData.Load(rData);
            // close that reader! 
            rData.Close();
            // close your connection to the spreadsheet! 
            oCon.Close();

            //时间原因只读了 Sheet1 如有需要请修改
            List<DataTable> dataTableList = new List<DataTable>();
            dataTableList.Add(dtYourData);

            return dataTableList;
        }
    }
}
