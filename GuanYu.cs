using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel
{
    public partial class GuanYu : Form
    {
        public GuanYu()
        {
            InitializeComponent();
        }

        private void GuanYu_Load(object sender, EventArgs e)
        {
            lblVersion.Text = "版本信息：2013.9.11测试版";
        }
    }
}
