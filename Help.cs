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
    public partial class FromHelp : Form
    {
        public FromHelp()
        {
            InitializeComponent();
        }

        private void Help_Load(object sender, EventArgs e)
        {
            


            textBox.Text = "1、导入的数据的页签（sheet）的名字必须为“SQL Results ”\r\n\r\n"
                + "2、导如的数据的顺序必须为：序号、流程ID、流程名称、环节ID、环节名称、可提交路径\r\n"
                + "建议SQL为：select ws.step_id as 环节ID ,ws.step_name as 环节名称 ,ws.step_path as 可提交路径"
                + "from workflow_step ws where ws.workflow_id in ('流程的ID') "
                + "order by ws.workflow_id,ws.step_id\r\n\r\n"
                + "3、目前此工具仅支持offic2003和offic2007，暂不支持wps\r\n\r\n"
                + "4、使用前请确认本机是否安装了Microsoft .NET Framework 4，如未安装请自行到微软官网下载并安装";


            textBox.SelectionStart = 0;
            textBox.SelectionLength = 0;
                
        }


    }
}
