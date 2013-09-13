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
    public partial class OperationExplain : Form
    {
        public OperationExplain()
        {
            InitializeComponent();
        }

        private void Help_Load(object sender, EventArgs e)
        {
            


            textBox.Text = "1、导入的数据的页签（sheet）的名字必须为“SQL Results ”\r\n\r\n"
                + "2、导如的数据的顺序必须为：序号、流程ID、流程名称、环节ID、环节名称、可提交路径、进一步处理、审批意见是否必填、开关编码\r\n\r\n"
                + "建议SQL为：select ws.workflow_id as 流程ID ,ws.workflow_name 流程名称,ws.step_id as 环节ID ,ws.step_name as 环节名称 ,ws.step_path as 可提交路径, ws.EXECUTE_FLAG as 进一步处理 ,ws.PUR_OPINION_WRITE as 审批意见是否必填 ,ws.SWITCHES as 开关编码"
                + "from workflow_step ws where ws.workflow_id in ('流程的ID')"
                + "order by ws.workflow_id,ws.step_id\r\n\r\n"
                + "3、目前此工具仅支持offic2003和offic2007，暂不支持wps\r\n\r\n"
                + "4、如果提示不能访问“Sheet1.xlsx”，请在进程管理器中将所有的excel.exe强制结束";





            textBox.SelectionStart = 0;
            textBox.SelectionLength = 0;
                
        }

    }
}
