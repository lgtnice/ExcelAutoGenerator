using System.Data;
using System.Text.RegularExpressions;
using System.IO;
using System;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace DataIO
{
    public class Methods
    {
        private static bool parseDyIsEmpty(string str) { return ((str == "") || (Regex.IsMatch(str, "F[0-9]{1,2}")) || (Regex.IsMatch(str, "Column[0-9]{1,2}"))); }
        
        public static void datatableExportToExcel(DataTable dt,string type,bool isExportColumnName,string prefix)
        {
            DataTable dt_r = new DataTable();
            if (isExportColumnName)
            {
                for (int i = 0; i < dt.Columns.Count; i++) { dt_r.Columns.Add(); }
                DataRow dr = dt_r.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++) { dr[i] = (parseDyIsEmpty(dt.Columns[i].ColumnName) ? "" : dt.Columns[i].ColumnName); }
                dt.Rows.Add(dr);
                for (int i = 0; i < dt.Rows.Count; i++) { dt_r.Rows.Add(dt.Rows[i].ItemArray); }
            }
            else { dt_r = dt; }
            string aimPath = Environment.CurrentDirectory + "\\生成文件\\";
            if (!Directory.Exists(aimPath)) { Directory.CreateDirectory(aimPath); }
            InteropExcel.Application application = new InteropExcel.ApplicationClass();
            application.DisplayAlerts = false;
            string fileName = prefix + ((DateTime.Now.Hour.ToString().Length == 1) ? ("0" + DateTime.Now.Hour.ToString()) : (DateTime.Now.Hour.ToString()))
                + ((DateTime.Now.Minute.ToString().Length == 1) ? ("0" + DateTime.Now.Minute.ToString()) : (DateTime.Now.Minute.ToString()))
                + ((DateTime.Now.Second.ToString().Length == 1) ? ("0" + DateTime.Now.Second.ToString()) : (DateTime.Now.Second.ToString())) + type;
            InteropExcel.Workbook workbook = application.Workbooks.Add(Type.Missing);
            switch (type)
            {
                case ".xlsx": workbook.SaveAs(aimPath + fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, InteropExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                case ".xls": workbook.SaveAs(aimPath + fileName, InteropExcel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, InteropExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                default: break;
            }
            InteropExcel.Worksheet worksheet = (InteropExcel.Worksheet)workbook.Worksheets[1];
            string[,] strs = new string[dt_r.Rows.Count, dt_r.Columns.Count];
            for (int i = 0; i < dt_r.Rows.Count; i++)
            {
                for (int j = 0; j < dt_r.Columns.Count; j++) { strs[i, j] = dt_r.Rows[i][j].ToString(); }
            }
            worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[dt.Rows.Count, dt.Columns.Count]).Value = strs;

            workbook.Save();
            workbook.Close(false, Type.Missing, Type.Missing);
            workbook = null;
            application.DisplayAlerts = true;
            application.Quit();
            application = null;
            GC.Collect();
        }

        public static void datatableExportToTXT(DataTable dt, bool isExportColumnName,string prefix)
        {
            DataTable dt_r = new DataTable();
            if (isExportColumnName)
            {
                for (int i = 0; i < dt.Columns.Count; i++) { dt_r.Columns.Add(); }
                DataRow dr = dt_r.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++) { dr[i] = (parseDyIsEmpty(dt.Columns[i].ColumnName) ? "" : dt.Columns[i].ColumnName); }
                dt.Rows.Add(dr);
                for (int i = 0; i < dt.Rows.Count; i++) { dt_r.Rows.Add(dt.Rows[i].ItemArray); }
            }
            else { dt_r = dt; }
            string aimPath = Environment.CurrentDirectory + "\\生成文件\\";
            if (!Directory.Exists(aimPath)) { Directory.CreateDirectory(aimPath); }
            string fileName = prefix + ((DateTime.Now.Hour.ToString().Length == 1) ? ("0" + DateTime.Now.Hour.ToString()) : (DateTime.Now.Hour.ToString()));
            fileName += ((DateTime.Now.Minute.ToString().Length == 1) ? ("0" + DateTime.Now.Minute.ToString()) : (DateTime.Now.Minute.ToString()));
            fileName += ((DateTime.Now.Second.ToString().Length == 1) ? ("0" + DateTime.Now.Second.ToString()) : (DateTime.Now.Second.ToString())) + ".txt";
            FileStream fs = File.Create(Path.Combine(aimPath,fileName));
            StreamWriter sw = new StreamWriter(fs);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++) { sw.Write(dt.Rows[i][j].ToString() + ","); }
                sw.WriteLine();
                sw.Flush();
            }
            sw.Flush();sw.Close();sw.Dispose();
            fs.Close();fs.Dispose();
        }
    }
}
