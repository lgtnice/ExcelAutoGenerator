using System;
using System.Windows.Forms;
using System.Collections;

namespace ZDYcontrol
{
    public partial class control_A : UserControl
    {
        public CheckBox cb
        {
            get { return checkBox_f; }
        }
        public ComboBox cbb
        {
            get { return comboBox1; }
        }
        public TextBox tb
        {
            get { return textBox1; }
        }
        public void reChange(bool ischecked,ArrayList comboboxContent,String textBoxString)
        {
            checkBox_f.Checked = ischecked; 
            if (comboboxContent != null) { foreach (object b in comboboxContent) { if (b is string) { comboBox1.Items.Add(b.ToString()); } } }
            if (textBoxString != null) { textBox1.Text = textBoxString; }
            if (comboBox1.Items.Count != 0) { comboBox1.SelectedIndex = 0; }
        }
        public control_A() { InitializeComponent();}
    }
}
