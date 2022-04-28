using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Text;

namespace DataIO
{
    public enum ExcelDataTableType
    {
        fourParts = 0,
        topBottom = 1,
        leftRight = 2,
    }
    public class ExcelDataTable : DataTable
    {
        private string excelName = "";
        private bool isFirstRowBeColumn = false;



        private DataTable firstTable = new DataTable();
        public DataTable FirstTable
        {
            get { return firstTable; }
            set { firstTable = value; }
        }
        private DataTable secondTable = new DataTable();
        public DataTable SecondTable
        {
            get { return secondTable; }
            set { secondTable = value; }
        }
        private DataTable thirdTable = new DataTable();
        public DataTable ThirdTable
        {
            get { return thirdTable; }
            set { thirdTable = value; }
        }
        private DataTable forthTable = new DataTable();
        public DataTable ForthTable
        {
            get { return forthTable; }
            set { forthTable = value; }
        }
        private DataTable topTable = new DataTable();
        public DataTable TopTable
        {
            get { return topTable; }
            set { topTable = value; }
        }
        private DataTable bottomTable = new DataTable();
        public DataTable BottomTable
        {
            get { return bottomTable; }
            set { bottomTable = value; }
        }
        private DataTable leftTable = new DataTable();
        public DataTable LeftTable
        {
            get { return leftTable; }
            set { leftTable = value; }
        }
        private DataTable rightTable = new DataTable();
        public DataTable RightTable
        {
            get { return rightTable; }
            set { rightTable = value; }
        }



        

        private ExcelDataTable() { }
        public ExcelDataTable(string path, string sheetName, bool isFirstRowToColumn,ExcelDataTableType excelType)
            : this()
        {
            this.excelName = Path.GetFileName(path);
            this.isFirstRowBeColumn = isFirstRowToColumn;
            getDataTable(path, sheetName, isFirstRowToColumn);
            parseMainTable(excelType);
        }
        public void parseMainTable(ExcelDataTableType excelType)
        {
            if (this == null) { return; }
            DataTable dt = new DataTable();
            if (isFirstRowBeColumn)
            {
                for (int i = 0; i < this.Columns.Count; i++) { dt.Columns.Add(); }
                DataRow dr = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++) { dr[i] = (parseDyIsEmpty(this.Columns[i].ColumnName) ? "" : this.Columns[i].ColumnName); }
                dt.Rows.Add(dr);
                for (int i = 0; i < this.Rows.Count; i++) { dt.Rows.Add(this.Rows[i].ItemArray); }
            }
            else { dt = this; }
            if (excelType == ExcelDataTableType.fourParts) 
            {
                if(!parseDyIsEmpty(dt.Rows[0][0].ToString())){ MessageBox.Show(string.Format("来源为{0}的table不合法",this.excelName)); return; }
                //计算secondTable的行数和列数
                int secondTable_ColumnCount = 1; int secondTable_RowCount = 1;
                while (parseDyIsEmpty(dt.Rows[0][secondTable_ColumnCount].ToString())) { secondTable_ColumnCount++; }
                while (true)
                {
                    bool flag = false;
                    for (int j_column = 0; j_column < secondTable_ColumnCount; j_column++)
                    { if (!parseDyIsEmpty(dt.Rows[secondTable_RowCount][j_column].ToString())) { flag = true; break; } }
                    if (flag) { break; }
                    secondTable_RowCount++;
                }
                //第二table过大，退出
                if ((dt.Rows.Count <= secondTable_RowCount) || (dt.Columns.Count <= secondTable_ColumnCount)) { return; }

                //计算出第一个firstDatatable
                for (int i = secondTable_ColumnCount; i < dt.Columns.Count; i++) { firstTable.Columns.Add(); }
                for (int i = 0; i < secondTable_RowCount; i++)
                {
                    DataRow dr = firstTable.NewRow();
                    for (int j = 0; j < firstTable.Columns.Count; j++) { dr[j] = dt.Rows[i][j + secondTable_ColumnCount].ToString(); }
                    firstTable.Rows.Add(dr);
                }
                //计算出第二个secondDatatable
                for (int i = 0; i < secondTable_ColumnCount; i++) { secondTable.Columns.Add(); }
                for (int i = 0; i < secondTable_RowCount; i++)
                {
                    DataRow dr = secondTable.NewRow();
                    for (int j = 0; j < secondTable.Columns.Count; j++) { dr[j] = dt.Rows[i][j].ToString(); }
                    secondTable.Rows.Add(dr);
                }
                //计算出第三个thridDatatable
                for (int i = 0; i < secondTable_ColumnCount; i++) { thirdTable.Columns.Add(); }
                for (int i = secondTable_RowCount; i < dt.Rows.Count; i++)
                {
                    DataRow dr = thirdTable.NewRow();
                    for (int j = 0; j < thirdTable.Columns.Count; j++) { dr[j] = dt.Rows[i][j].ToString(); }
                    thirdTable.Rows.Add(dr);
                }
                //计算出第四个forthDatatable
                for (int i = secondTable_ColumnCount; i < dt.Columns.Count; i++) { forthTable.Columns.Add(); }
                for (int i = secondTable_RowCount; i < dt.Rows.Count; i++)
                {
                    DataRow dr = forthTable.NewRow();
                    for (int j = 0; j < forthTable.Columns.Count; j++) { dr[j] = dt.Rows[i][j + secondTable_ColumnCount].ToString(); }
                    forthTable.Rows.Add(dr);
                }
            }
            else if (excelType == ExcelDataTableType.topBottom)
            {
                for (int i = 0; i < dt.Columns.Count; i++) { topTable.Columns.Add(); bottomTable.Columns.Add(); }
                int BottomTableRowsCount = parse_BottomTableRowsCount(dt);
                int TopTableRowsCount = dt.Rows.Count - BottomTableRowsCount;
                for (int i = 0; i < (dt.Rows.Count - BottomTableRowsCount); i++) { topTable.NewRow(); topTable.Rows.Add(dt.Rows[i].ItemArray); }
                for (int i = 0; i < (BottomTableRowsCount); i++) { bottomTable.NewRow(); bottomTable.Rows.Add(dt.Rows[i + TopTableRowsCount].ItemArray); }
            }
            else if (excelType == ExcelDataTableType.leftRight)
            {
                for (int i = 0; i < parse_LeftTableColumnCount(dt); i++) { leftTable.Columns.Add(); }
                for (int i = 0; i < dt.Rows.Count; i++)
                { leftTable.Rows.Add(); for (int j = 0; j < parse_LeftTableColumnCount(dt); j++) { leftTable.Rows[i][j] = dt.Rows[i][j].ToString(); } }
                for (int i = 0; i < dt.Columns.Count - parse_LeftTableColumnCount(dt); i++) { rightTable.Columns.Add(); }
                for (int i = 0; i < dt.Rows.Count; i++)
                { rightTable.Rows.Add(); for (int j = parse_LeftTableColumnCount(dt); j < dt.Columns.Count; j++) { rightTable.Rows[i][j - parse_LeftTableColumnCount(dt)] = dt.Rows[i][j].ToString(); } }
            }
            else { MessageBox.Show("ExcelDataTableType不合法"); }
        }
        public void getDataTable(string path, string sheetName, bool isFirstRowToColumn)
        {
            if (Path.IsPathRooted(path)) { if (!File.Exists(path)) { return; } }
            else { path = Path.Combine(Environment.CurrentDirectory, path); if (!File.Exists(path)) { return; } }
            switch (Path.GetExtension(path))
            {
                case ".csv":
                    {
                        FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                        StreamReader sr = new StreamReader(fs, Encoding.Default);
                        string[] columnName = null; string[] normalRow = null; string strLine = null;
                        strLine = sr.ReadLine();
                        if (isFirstRowToColumn)
                        {
                            isFirstRowToColumn = false;
                            columnName = strLine.Split(',');
                            for (int i = 0; i < columnName.Length; i++)
                            {
                                DataColumn dc = new DataColumn(columnName[i]);
                                this.Columns.Add(dc);
                            }
                        }
                        else
                        {
                            normalRow = strLine.Split(',');
                            for (int i = 0; i < normalRow.Length; i++)
                            {
                                DataColumn dc = new DataColumn();
                                this.Columns.Add(dc);
                            }
                            DataRow dr = this.NewRow();
                            for (int j = 0; j < normalRow.Length; j++) { dr[j] = normalRow[j]; }
                            this.Rows.Add(dr);
                        }
                        while ((strLine = sr.ReadLine()) != null)
                        {
                            normalRow = strLine.Split(',');
                            DataRow dr = this.NewRow();
                            for (int j = 0; j < normalRow.Length; j++) { dr[j] = normalRow[j]; }
                            this.Rows.Add(dr);
                        }
                        fs.Flush(); fs.Close();
                        break;
                    }
                case ".xlsx":
                case ".xls":
                    {
                        string conString = "";
                        if (Path.GetExtension(path) == ".xlsx")
                        {
                            conString = string.Format("Provider=Microsoft.Ace.OleDb.12.0;data source={0};Extended Properties='Excel 12.0;HDR={1};IMEX=1;'", path, ((isFirstRowToColumn) ? "Yes" : "No"));
                        }
                        else
                        {
                            conString = string.Format("Provider=Microsoft.Ace.OleDb.12.0;data source={0};Extended Properties='Excel 12.0;HDR={1};IMEX=1;'", path, ((isFirstRowToColumn) ? "Yes" : "No"));
                            //conString = string.Format("Provider=Microsoft.Jet.OleDb.4.0;data source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1;'", path, ((isFirstRowToColumn) ? "Yes" : "No"));
                        }
                        OleDbConnection excelCon = new OleDbConnection(conString);
                        excelCon.Open();
                        if ((sheetName != "") && (sheetName != null))
                        {
                            bool flag = true;
                            foreach (DataRow row in excelCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows)
                            { if ((string)row["TABLE_NAME"] == sheetName) { flag = false; break; } }
                            if (flag) { MessageBox.Show("未找到此名称的sheet"); return; }
                        }
                        else { sheetName = excelCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString(); }
                        string cmdText = string.Format("select * FROM [Excel 12.0;HDR={2};DATABASE={0}].[{1}]", path, sheetName, ((isFirstRowToColumn) ? "Yes" : "No"));
                        OleDbDataAdapter oda = new OleDbDataAdapter(cmdText, excelCon); 
                        oda.Fill(this);
                        excelCon.Dispose(); oda.Dispose();
                        break;
                    }
                default: { MessageBox.Show(string.Format("{0}文件后缀名不合法", Path.GetExtension(path))); break; }
            }
        }
        private bool parseDyIsEmpty(string str) { return ( (str == "") || (Regex.IsMatch(str, "F[0-9]{1,2}")) || (Regex.IsMatch(str, "Column[0-9]{1,2}")) ); }
        private int parse_BottomTableRowsCount(DataTable dt)
        {
            int rowCount = 0; int x; int y;
            for (int i = 0; i < 5; i++)
            {
                if (dt.Rows[i][0].ToString() == dt.Rows[i + 1][0].ToString()) 
                { rowCount = dt.Rows.Count - i; break; }
                if ((int.TryParse(dt.Rows[i][0].ToString(),out x)) && (int.TryParse(dt.Rows[i + 1][0].ToString(),out y)))
                { if ((x + 1) == y) { rowCount = dt.Rows.Count - i; break; } }
            }
            return rowCount;
        }
        private int parse_LeftTableColumnCount(DataTable dt)
        {
            int j = 0;
            for (; j < 5; j++)
            {
                int count = 0;
                for (int i = 0; i < 5; i++) {if (dt.Rows[i][j].ToString() == "") { count++; } }
                if (count > 2) { break; }
            }
            return j;
        }
    }
}
