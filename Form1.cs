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
                dataGridView.DataSource = ExcelHelp.LoadDataFromExcel(Openlujing).Tables[0];
            }

            if (ExcelHelp.updateExcel(dataGridView) != null)
            {
                //更新状态为数据转换中
                lblState.Text = "状态：数据转换中";

                //数据转换
                dataGridView.DataSource = ExcelHelp.updateExcel(dataGridView);

                //更新状态
                lblState.Text = "状态：数据转换完成，待导出";
                //获取导出路径
                string lujing = ExcelHelp.SaveFileDialog(saveFileDialog);
                //展示导出路径
                lblSave.Text = lblSave.Text.ToString() + lujing;
                //更新状态
                lblState.Text = "状态：导出中，数据量较大，请稍后";
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

                if (ExcelHelp.updateExcel(dataGridView) != null)
                {
                    lblState.Text = "状态：数据转换中";
                    //数据转换
                    dataGridView.DataSource = ExcelHelp.updateExcel(dataGridView);
                    //更新状态
                    lblState.Text = "状态：数据转换完成";
                }
                else
                {
                    MessageBox.Show("转换失败");
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


        private void tsmiGuanyu_Click(object sender, EventArgs e)
        {
            GuanYu guanyu = new GuanYu();
            guanyu.ShowDialog();
        }

        private void tmsiHelp_Click(object sender, EventArgs e)
        {
            FromHelp help = new FromHelp();
            help.ShowDialog();
        }

        private void tsmiexplain_Click(object sender, EventArgs e)
        {
            MessageBox.Show("使用前请先查看帮助");
        }



    }
}
