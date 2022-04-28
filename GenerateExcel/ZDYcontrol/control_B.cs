using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ZDYcontrol
{
    public partial class control_B : UserControl
    {
        public Button bt
        {
            get { return button1; }
        }
        public TextBox tb
        {
            get { return textBox1; }
        }
        public control_B()
        {
            InitializeComponent();
        }
    }
}
