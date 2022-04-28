using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Windows.Forms;
using ZDYcontrol;
using DataIO;
using System.Text;


namespace WindowsFormsApplication1
{
    public partial class Form_Main : Form
    {
        public Form_Main()
        {
            InitializeComponent();
            initialForm_MES();
        }



        //处理字符串
        private string handleString(string s ,string flag)
        {
            switch (flag)
            {
                case "1": 
                    { return s; }
                case "2":
                    {
                        char[] c = s.ToCharArray();
                        if (c.Length == 12)
                        {
                            return ("" + c[0] + c[1] + ":" + c[2] + c[3] + ":" + c[4] + c[5] + ":" + c[6] + c[7] + ":" + c[8] + c[9] + ":" + c[10] + c[11]);
                        }
                        else if (c.Length == 17)
                        {
                            return ("" + c[0] + c[1] + ":" + c[3] + c[4] + ":" + c[6] + c[7] + ":" + c[9] + c[10] + ":" + c[12] + c[13] + ":" + c[15] + c[16]);
                        }
                        else
                        {
                            return "";
                        }
                    }
                case "3":
                    {
                        return s + " " + "-n";
                    }
                case "4":
                    {
                        char[] c = s.ToCharArray();
                        if(s.Length == 12){return s;}
                        else if(s.Length == 17){return ("" + c[0] + c[1] + c[3] + c[4] + c[6] + c[7] + c[9] + c[10] + c[12] + c[13] + c[15] + c[16]);}
                        else{return "";}
                    }
                case "5":
                    {
                        string[] ss = s.Split('-');
                        return ss[1];
                    }
                case "6":
                    {
                        char[] c = s.ToCharArray();
                        string dd = "" + c[c.Length - 8] + c[c.Length - 7] + c[c.Length - 6] +
                        c[c.Length - 5] + c[c.Length - 4] + c[c.Length - 3] +
                        c[c.Length - 2] + c[c.Length - 1];
                        return dd;
                    }
                case "7":
                    {
                        char[] c = s.ToCharArray();
                        string dd = "" + c[c.Length - 9] + c[c.Length - 8] + c[c.Length - 7] + c[c.Length - 6] +
                        c[c.Length - 5] + c[c.Length - 4] + c[c.Length - 3] +
                        c[c.Length - 2] + c[c.Length - 1];
                        return dd;
                    }
                case "8":
                    {
                        string[] ss = s.Split('-');
                        return ss[0];
                    }
                default:
                    return "";
            }
        }



        //上传到MES的数据生成，初始化界面按钮
        private void initialForm_MES()
        {
            ExcelDataTable dt_config = new ExcelDataTable("config_MES.xls", "", true, ExcelDataTableType.fourParts);
            ArrayList tempAL_config = new ArrayList();
            int rowNoHyConfig = 0;
            for (; rowNoHyConfig < dt_config.ThirdTable.Rows.Count; rowNoHyConfig++)
            {
                if (dt_config.ThirdTable.Rows[rowNoHyConfig][0].ToString() == "自己选择") { break; }
            }
            for (int i = 0; i < dt_config.FirstTable.Columns.Count; i++)
            {
                if (dt_config.ForthTable.Rows[rowNoHyConfig][i].ToString() != "")
                {
                    tempAL_config.Add(dt_config.FirstTable.Rows[0][i].ToString());
                }
            }
            foreach (Control item in this.flowLayoutPanel1.Controls)
            {
                if (!(item is control_A)) { continue; }
                control_A y = (control_A)item;
                if (tempAL_config.Count > 0)
                {
                    ArrayList al = new ArrayList();
                    al.Add(tempAL_config[0].ToString());
                    y.reChange(true, al, "");
                    tempAL_config.RemoveAt(0);
                }
                else { break; }
            }
            dt_config.Dispose();
        }
        //上传到MES的数据生成，开始
        private void button_MesEnter_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            DataTable dt = Datatable_MES();
            if ((dt == null) || (dt.Rows.Count == 0)) { MessageBox.Show("对上传到MES的DataTable，生成失败"); return; }
            DataIO.Methods.datatableExportToExcel(dt, ".xls", false, "上传到MES的表");
            dt.Dispose();
            this.WindowState = FormWindowState.Normal;
        }
        //上传到MES的数据生成，获取CSV
        private void button_MesSourcePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "支持的文件格式|*.xls;*.xlsx;*.csv";
            //op.Filter = "支持的文件格式|*.csv";
            op.ShowDialog();
            textBox_MesSourcePath.Text = op.FileName;
        }
        //上传到MES的数据生成，datatable
        private DataTable Datatable_MES()
        {
            if (!File.Exists(textBox_MesSourcePath.Text)) { MessageBox.Show("源数据文件的路径下找不到有效文件"); return null; }
            ExcelDataTable dt_source = new ExcelDataTable(textBox_MesSourcePath.Text, "", false,ExcelDataTableType.topBottom);
            ExcelDataTable dt_config = new ExcelDataTable("config_MES.xls","", false,ExcelDataTableType.fourParts);
            DataTable dt = dt_config.FirstTable.Copy();
            if (dt_config.FirstTable.Rows.Count == 0) { MessageBox.Show("MES配置文件出错"); return dt; }
            if (dt_source.BottomTable.Rows.Count == 0) { MessageBox.Show("MES源文件出错"); return dt; }
            for (int tempi = 0; tempi < dt_source.BottomTable.Rows.Count; tempi++) { dt.Rows.Add(new object[] { }); }
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                string item = null; int j = 0;
                for (; j < dt_config.ForthTable.Rows.Count; j++)
                {
                    if (dt_config.ForthTable.Rows[j][i].ToString() != "")
                    { item = dt_config.ForthTable.Rows[j][i].ToString(); break; }
                }
                if (item == null) { continue; }
                else
	            {
                    if (dt_config.ThirdTable.Rows[j][0].ToString() == "自己选择")
                    {
                        string aimStr = "";
                        foreach (Control con in this.flowLayoutPanel1.Controls)
                        {
                            if (!(con is control_A)) { continue; }
                            control_A conA = (control_A)con;
                            if ((conA.cb.Checked == true) && (conA.cbb.Text == dt.Rows[0][i].ToString())) { aimStr = conA.tb.Text; break; }
                        }
                        for (int d_row = dt_config.FirstTable.Rows.Count; d_row < dt_config.FirstTable.Rows.Count + dt_source.BottomTable.Rows.Count; d_row++)
                        { dt.Rows[d_row][i] = aimStr; }
                    }
                    else
                    {
                        int sour_j = 0;
                        for (; sour_j < dt_source.TopTable.Columns.Count; sour_j++)
                        {
                            if (dt_source.TopTable.Rows[0][sour_j].ToString() == dt_config.ThirdTable.Rows[j][0].ToString()) { break; }
                        }
                        if (sour_j == dt_source.TopTable.Columns.Count)
                        { MessageBox.Show(string.Format("在MES源数据的表内没有找到此列名：{0}", dt_config.ThirdTable.Rows[j][0].ToString())); break; }
                        for (int d_row = 0; d_row < dt_source.BottomTable.Rows.Count; d_row++)
                        { dt.Rows[d_row + dt_config.FirstTable.Rows.Count][i] = handleString(dt_source.BottomTable.Rows[d_row][sour_j].ToString(),item) ; }
                    }
	            }            
            }
            return dt;
        }
        //生成生产数据
        private void button_sc_Click(object sender, EventArgs e)
        {
            ExcelDataTable dt_config = new ExcelDataTable("config_MES.xls", "", false, ExcelDataTableType.fourParts);
        }


        //产测软件用的TXT生成,获取CSV
        private void button_CCRJSourcePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "支持的文件格式|*.xls;*.xlsx;*.csv";
            //op.Filter = "支持的文件格式|*.csv";
            op.ShowDialog();
            textBox_CCRJSourcePath.Text = op.FileName;
        }
        //产测软件用的TXT生成,开始
        private void button_CCRJEnter_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            DataTable dt = Datatable_CCRJ();
            if ((dt == null) || (dt.Rows.Count == 0)) { MessageBox.Show("对产测软件用表的DataTable，生成失败"); return; }
            DataIO.Methods.datatableExportToTXT(dt, false, "产测软件的用表");
            dt.Dispose();
            this.WindowState = FormWindowState.Normal;
        }
        //产测软件用的TXT生成,datatable
        private DataTable Datatable_CCRJ()
        {
            if (!File.Exists(textBox_CCRJSourcePath.Text)) { MessageBox.Show("产测软件的源数据文件路径下找不到有效文件"); return null; }
            ExcelDataTable dt_source = new ExcelDataTable(textBox_CCRJSourcePath.Text, "", false,ExcelDataTableType.topBottom);
            ExcelDataTable dt_config = new ExcelDataTable("config_CCRJ.xls", "", false,ExcelDataTableType.leftRight);
            DataTable dt = new DataTable();
            for (int i = 0; i < dt_config.RightTable.Columns.Count; i++) { dt.Columns.Add(); }
            for (int i = 0; i < dt_source.BottomTable.Rows.Count; i++) { dt.Rows.Add(); }
            for (int j = 0; j < dt_config.RightTable.Columns.Count; j++)
            {
                string item = null; string flag = null;
                for (int config_i = 0; config_i < dt_config.RightTable.Rows.Count; config_i++)
                {
                    if (dt_config.RightTable.Rows[config_i][j].ToString() != "")
                    { item = dt_config.LeftTable.Rows[config_i][0].ToString(); flag = dt_config.RightTable.Rows[config_i][j].ToString(); break; }
                }
                if ((flag == "") || (flag == null)) { continue; }
                else
                {
                    int source_j = 0;
                    for (; source_j < dt_source.TopTable.Columns.Count; source_j++) { if (dt_source.TopTable.Rows[0][source_j].ToString() == item) { break; } }
                    if ((source_j < 0) || (source_j >= dt_source.TopTable.Columns.Count)) { MessageBox.Show(string.Format("在打开的源文件里面没有发现名称为{0}的列", item)); break; }
                    for (int i = 0; i < dt_source.BottomTable.Rows.Count; i++) { dt.Rows[i][j] = handleString(dt_source.BottomTable.Rows[i][source_j].ToString(),flag); }
                }
            }
            return dt;
        }

        

    }
}
