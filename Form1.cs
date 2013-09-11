using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Web;



namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //标识是否数据转换过。0表示没转换，1表示转换过
        int zhuanhuan ;
        String Openlujing;


        /// <summary>
        /// 单击导入Excel按钮的事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInput_Click(object sender, EventArgs e)
        {
            zhuanhuan = 0;

            //string lujing = "F:\\vs2010\\Excel\\测试.xls";
            Openlujing = ExcelHelp.OpenFileDialog(openFileDialog);

            //判断路径是否为空
            if (Openlujing == null || Openlujing.Equals(null))
            {
                MessageBox.Show("没有选择Excel文件！无法进行数据导入");
            }
            else
            {
                //设置路径的位置
                lblOpen.Text = lblOpen.Text.ToString() + Openlujing;
                //更新状态
                lblState.Text = "状态：数据导入中，请稍后";
                dataGridView.DataSource = null;
                //LoadDataFromExcel(lujing);
                //MessageBox.Show("文件路径为：" + lujing);
                //为dataGridView指定数据源"SQL Results$"


                dataGridView.DataSource = ExcelHelp.LoadDataFromExcel(Openlujing).Tables[0];
            
                //设置dataGridView为不可排序模式
                ExcelHelp.ForbidSortColumn(dataGridView);

                //更新状态为数据转换中
                lblState.Text = "状态：数据导入完成";
                
            }

          


            
        }

        /// <summary>
        /// 单击导出Excel按钮的事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            if(zhuanhuan==1)
            {
                dataGridView.DataSource = null;
                dataGridView.DataSource = ExcelHelp.LoadDataFromExcel(Openlujing).Tables[0];
            }

            //更新状态为数据转换中
            lblState.Text = "状态：数据转换中";
            //数据转换
            dataGridView.DataSource = ExcelHelp.updateExcel(dataGridView);

            if (dataGridView.DataSource != null)
            {
                

                //更新状态
                lblState.Text = "状态：数据转换完成，待导出";
                //获取导出路径
                string lujing = ExcelHelp.SaveFileDialog(saveFileDialog);
                //展示导出路径
                lblSave.Text = lblSave.Text.ToString() + lujing;
                //更新状态
                lblState.Text = "状态：导出中，数据量较大，请稍后";
                MessageBox.Show("因数据量较大，所以导出时间可能较长，请耐心等待\r\n\r\n点击【确定】按钮后数据开始导出");
                ExcelHelp.SaveDataTableToExcel((System.Data.DataTable)this.dataGridView.DataSource, lujing);
                //更新状态
                lblState.Text = "状态：数据导出完成";


                
            }
            else
            {
                MessageBox.Show("转换失败");
            }

           
        }

        private void btncs_Click(object sender, EventArgs e)
        {
            if (zhuanhuan == 0)
            {
                zhuanhuan = 1;

                lblState.Text = "状态：数据转换中";
                //数据转换
                dataGridView.DataSource = ExcelHelp.updateExcel(dataGridView);

                if (dataGridView.DataSource != null)
                {
                    //更新状态
                    lblState.Text = "状态：数据转换完成";
                    MessageBox.Show("数据转换完成");
                }
                else
                {
                    MessageBox.Show("数据转换失败");
                }
            }
            else
            {
                MessageBox.Show("数据以转换过，无法再次转换");
            }
            
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            lblexplain.Text = "说明：使用前请先查看帮助";
        }


        


        

        private void btncs2_Click(object sender, EventArgs e)
        {
            //string lujing = "D:\\多个页签测试.xls";
            //cscs.SaveDataTableToExcel((System.Data.DataTable)this.dataGridView.DataSource, lujing);
        }


        /// <summary>
        /// 说明菜单的单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmiexplain_Click(object sender, EventArgs e)
        {
            MessageBox.Show("使用前请先查看帮助");
        }



        /// <summary>
        /// 关于的单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmiGuanyu_Click(object sender, EventArgs e)
        {
            GuanYu guanyu = new GuanYu();
            guanyu.ShowDialog();
        }


        /// <summary>
        /// 操作说明的单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmiOperationExplain_Click_1(object sender, EventArgs e)
        {
            OperationExplain operationExplain = new OperationExplain();
            operationExplain.ShowDialog();
        }


        /// <summary>
        /// 更新说明的单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsmiUpdateExplain_Click(object sender, EventArgs e)
        {
            MessageBox.Show("新增功能：\r\n"
                + "1、增加了对进一步处理的支持\r\n"
                + "2、增加了对审批意见是否必填的支持\r\n"
                + "\r\n"
                + "更新功能：\r\n"
                + "1、更新了可提交路径的替换逻辑\r\n"
                + "2、修正了部分BUG\r\n"
                + "3、更改了菜单栏的排序和说明\r\n"
                );
        }

    }
}
