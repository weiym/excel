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

            //此处需要判断数据是否转换过，如果转换过需要重新加载数据，否则直接转换数据会有问题。
            if(zhuanhuan==1)
            {
                dataGridView.DataSource = null;
                dataGridView.DataSource = ExcelHelp.LoadDataFromExcel(Openlujing).Tables[0];
            }

            //更新状态为数据转换中
            lblState.Text = "状态：数据转换中";
            //数据转换
            dataGridView.DataSource = ExcelHelp.updateExcel(dataGridView);
            //是否转换的标示改为1
            zhuanhuan = 1;
            //更新dataGridView的颜色
            ExcelHelp.updateDataGridViewColor(dataGridView);

            if (dataGridView.DataSource != null)
            {
                

                //更新状态
                lblState.Text = "状态：数据转换完成，待导出";
                //获取导出路径
                string lujing = ExcelHelp.SaveFileDialog(saveFileDialog);

                if (lujing == null)
                {
                    MessageBox.Show("您未选择文件保存的位置和名称，数据无法导出，请重试");

                }
                else
                {
                    //展示导出路径
                    lblSave.Text = lblSave.Text.ToString() + lujing;
                    //更新状态
                    lblState.Text = "状态：导出中，数据量较大，请稍后";
                    MessageBox.Show("因数据量较大，所以导出时间可能较长，请耐心等待\r\n\r\n点击【确定】按钮后数据开始导出");
                    ExcelHelp.SaveDataTableToExcel((System.Data.DataTable)this.dataGridView.DataSource, lujing);
                    //更新状态
                    lblState.Text = "状态：数据导出完成";
 
                }
                


                
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
                    //更新dataGridView的颜色
                    ExcelHelp.updateDataGridViewColor(dataGridView);
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


            //MessageBox.Show("居中设置");
            //System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            //dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            //this.dataGridView.DefaultCellStyle = dataGridViewCellStyle1; 
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
                + "3、增加展示时每个流程一种颜色\r\n"
                + "4、增加了展示时对相同的数据合并单元格\r\n"
                + "\r\n"
                + "更新功能：\r\n"
                + "1、更新了可提交路径的替换逻辑\r\n"
                + "2、修正了部分BUG\r\n"
                + "3、更改了菜单栏的排序和说明\r\n"
                );
        }



   
        /// <summary>
        /// 网上拷贝的代码，不懂什么意思，用于单元格合并
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView_CellPainting_1(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // 循环  i代表了要合并的列，此处固化为了只有5列，后续如果需要可以改成动态的
            for (int i = 1; i < 5;i++ )
            {
                //MessageBox.Show("重绘单元格");
                // 对第1列相同单元格进行合并
                if (e.ColumnIndex == i && e.RowIndex != -1)
                {
                    using (Brush gridBrush = new SolidBrush(this.dataGridView.GridColor), backColorBrush = new SolidBrush(e.CellStyle.BackColor))
                    {
                        using (Pen gridLinePen = new Pen(gridBrush))
                        {
                            // 清除单元格
                            e.Graphics.FillRectangle(backColorBrush, e.CellBounds);

                            // 画 Grid 边线（仅画单元格的底边线和右边线）
                            // 如果下一行和当前行的数据不同，则在当前的单元格画一条底边线
                            if (e.RowIndex < dataGridView.Rows.Count - 1 &&
                            dataGridView.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value.ToString() !=
                            e.Value.ToString())

                                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left,
                                e.CellBounds.Bottom - 1, e.CellBounds.Right - 1,
                                e.CellBounds.Bottom - 1);
                            // 画右边线
                            e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1,
                            e.CellBounds.Top, e.CellBounds.Right - 1,
                            e.CellBounds.Bottom);

                            // 画（填写）单元格内容，相同的内容的单元格只填写第一个
                            if (e.Value != null)
                            {
                                //当前行的数据大于0，并且上一行的数据和当前行的数据相同
                                if (e.RowIndex > 0 &&
                                dataGridView.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() ==
                                e.Value.ToString())
                                { }
                                else
                                {
                                    e.Graphics.DrawString((String)e.Value, e.CellStyle.Font,
                                        Brushes.Black, e.CellBounds.X + 2,
                                        e.CellBounds.Y + 5, StringFormat.GenericDefault);
                                }
                            }
                            e.Handled = true;
                        }
                    }
                }
            }

           
        }




    }
}
