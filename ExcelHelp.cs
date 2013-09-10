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
using System.Reflection;
using System.Diagnostics;


namespace Excel
{
    public class ExcelHelp
    {

        public static DateTime beforeTime;			//Excel启动之前时间
        public static DateTime afterTime;				//Excel启动之后时间

        public static void cs()
        {
            String a = "001002";
            String b = null;
            b = a.Replace("001", "ceshi");

            MessageBox.Show(b);

        }



        /// <summary>
        /// 更新DataGridView的值，并返回一个dataTable。
        /// </summary>
        /// <param name="dataGridView">传入一个DataGridView</param>
        /// <returns>返回一个dataTable</returns>
        public static System.Data.DataTable updateExcel(DataGridView dataGridView)
        {

            int i = 0;
            int j = 0;
            int k = 0;
            //String before = null;
            //String after = null;
            //计数器，用于在某些时候记录当前已写入的dataGridView的行数
            int counter = 0;
            //可提交路径
            String lujing = null;
            //用于拆分可提交路径
            String[] lujingshuzu = new String[] { };



            //定义一个dataTable
            System.Data.DataTable dt = new System.Data.DataTable();


            //定义一个dataTable第一列为【序号】
            DataColumn dc1 = new DataColumn("序号");
            //定义一个dataTable第二列为【流程ID】
            DataColumn dc2 = new DataColumn("流程ID");
            //定义一个dataTable第三列为【流程名称】
            DataColumn dc3 = new DataColumn("流程名称");
            //定义一个dataTable第四列为【环节ID】
            DataColumn dc4 = new DataColumn("环节ID");
            //定义一个dataTable第五列为【环节名称】
            DataColumn dc5 = new DataColumn("环节名称");
            //定义一个dataTable第六列为【可提交路径】
            DataColumn dc6 = new DataColumn("可提交路径");

            //将定义的列放到datatable中
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            dt.Columns.Add(dc3);
            dt.Columns.Add(dc4);
            dt.Columns.Add(dc5);
            dt.Columns.Add(dc6);


            try
            {

                #region  此段代码作用是将可提交路径根据【；】进行分割换行

                //二维数组，用于将可提交路径换行，数组的行数动态生成，列数固定为6
                String[,] array = new String[dataGridView.Rows.Count, 6];

                //循环将dataGridView的数值赋到二维数组中
                for (int a = 0; a < dataGridView.Rows.Count; a++)
                {

                    for (int b = 0; b < 6; b++)
                    {
                        lujing = dataGridView.Rows[a].Cells[b].Value.ToString();

                        #region  删除部分特殊字符
                        //删除#A#
                        lujing = lujing.Replace("#A#", "");
                        //删除#N#
                        lujing = lujing.Replace("#N#", "");
                        //删除#T#
                        lujing = lujing.Replace("#T#", "");
                        //删除WORKFLOWSECRSELE
                        lujing = lujing.Replace("WORKFLOWSECRSELE", "");
                        //删除SECRNEXTSTEP
                        lujing = lujing.Replace("SECRNEXTSTEP", "");
                        #endregion


                        //将删除过部分特殊字符的dataGridView的数据动态赋给 二维数组
                        array[a, b] = lujing;
                    }
                }

                //临时测试，展示dataGridView的下标为2的行，第三列的数据
                //MessageBox.Show(array[2,3]);


                //定义一个集合
                List<String> counters = new List<string>();
                //获取流程的数量
                counters = cordysNumber(dataGridView);

                if (counters == null)
                {
                    MessageBox.Show("未找到需要转换的流程");
                    return null;
                }

                //根据数组的长度定义循环的次数
                for (int g = 0; g < counters.Count; g++)
                {
                    //因为第一次和后续几次的取下标的逻辑不同，所以需要分开判断   
                    if(g==0)
                    {
                        //判断数组第一个值是否大于0
                        if ( Int32.Parse(counters[0])>0)
                        {
                            #region 进行路径拆分
                            //循环将路径那一列并进行拆分
                            for (int c = 0; c < Int32.Parse(counters[0])+1; c++)
                            {
                                //MessageBox.Show(array[c,3]);

                                /**
                                 * 将路径的那一列，根据“；”进行拆分,    StringSplitOptions.RemoveEmptyEntries的作用为去除空值
                                 * 因为999和并签的部分环节对应的可提交路径为空，所以此处会把999和并签的部分环节给去掉,
                                 * */
                                lujingshuzu = array[c, 5].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                                //将路径的那一列，根据“；”进行拆分。   因为999的可提交路径为空，所以不能在此去除空值，否则会导致999无法被写入，无法被转换
                                //lujingshuzu = array[c, 3].Split(new char[] { ';' });




                                #region  此段代码作用是二维数组的数据做一定处理后，赋给datatable

                                /**
                                 * 根据数据内有没有值，也就是数组的长度不同，有不同的逻辑控制
                                 **/


                                //如果数组内没有值，也就是数组的长度为0
                                if (lujingshuzu.Length == 0)
                                {
                                    //定义一个dataTable的行
                                    DataRow dr = dt.NewRow();
                                    //为dataTable增加一行
                                    dt.Rows.Add(dr);

                                    //给datatable的第counter行的6列赋值
                                    dt.Rows[counter][0] = counter + 1;//序号
                                    dt.Rows[counter][1] = array[c, 1];//流程标识
                                    dt.Rows[counter][2] = array[c, 2];//流程名称
                                    dt.Rows[counter][3] = array[c, 3];//环节ID
                                    dt.Rows[counter][4] = array[c, 4];//环节名称
                                    dt.Rows[counter][5] = "";//可提交路径



                                    //计数器加上数组的长度，用于从第N行重新开始赋值
                                    counter = counter + 1;

                                }
                                else//数组内有值，也就是数组的长度大于0
                                {
                                    //循环将数组的值赋给datatable
                                    for (int d = 0; d < lujingshuzu.Length; d++)
                                    {
                                        //定义一个dataTable的行
                                        DataRow dr = dt.NewRow();
                                        //为dataTable增加一行
                                        dt.Rows.Add(dr);


                                        //dataGridView.Rows.Add();
                                        //定义一个临时计数器
                                        int counterls = counter + d;
                                        //临时测试，展示C的值
                                        //MessageBox.Show("C的值为" + c );


                                        //临时测试，展示lujingshuzu的值
                                        //for (int e = 0; e < lujingshuzu.Length;e++ )
                                        //{
                                        //    MessageBox.Show(lujingshuzu[e]);
                                        //}


                                        //因为第一次的逻辑和之后的不一样，所以在此判断是否是第一次，主要是序号的逻辑不通
                                        if (c == 0)
                                        {
                                            //判断根据；拆分后的值是否为空，不为空则赋值，如果为空则跳出循环
                                            if (lujingshuzu[d] != null && lujingshuzu[d] != "")
                                            {
                                                //dataGridView.Rows[d].Cells[1].Value = array[c, 1];
                                                //dataGridView.Rows[d].Cells[2].Value = array[c, 2];   
                                                //dataGridView.Rows[d].Cells[3].Value = lujingshuzu[d];

                                                ////给datatable的第N行的6列赋值
                                                dt.Rows[counterls][0] = d + 1;//序号
                                                dt.Rows[counterls][1] = array[c, 1];//流程标识
                                                dt.Rows[counterls][2] = array[c, 2];//流程名称
                                                dt.Rows[counterls][3] = array[c, 3];//环节ID
                                                dt.Rows[counterls][4] = array[c, 4];//环节名称
                                                dt.Rows[counterls][5] = lujingshuzu[d];//可提交路径

                                                //临时测试，展示拆分后的数据
                                                //MessageBox.Show("第C" + c + "次的数据为：" + dataGridView.Rows[d].Cells[3].Value.ToString());
                                            }
                                            else
                                            {
                                                continue;
                                            }

                                        }
                                        else
                                        {
                                            if (lujingshuzu[d] != null && lujingshuzu[d] != "")
                                            {

                                                //dataGridView.Rows[counterls].Cells[1].Value = array[c, 1];
                                                //dataGridView.Rows[counterls].Cells[2].Value = array[c, 2];
                                                //dataGridView.Rows[counterls].Cells[3].Value = lujingshuzu[d];

                                                ////给datatable的第N行的6列赋值
                                                dt.Rows[counterls][0] = counterls + 1;//序号
                                                dt.Rows[counterls][1] = array[c, 1];//流程标识
                                                dt.Rows[counterls][2] = array[c, 2];//流程名称
                                                dt.Rows[counterls][3] = array[c, 3];//环节ID
                                                dt.Rows[counterls][4] = array[c, 4];//环节名称
                                                dt.Rows[counterls][5] = lujingshuzu[d];//可提交路径
                                                //临时测试，展示拆分后的数据
                                                //MessageBox.Show("第counterls" + c + "次的数据为：" + dataGridView.Rows[counterls].Cells[3].Value.ToString());
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }

                                        //临时测试，展示拆分后的数据
                                        //MessageBox.Show("第"+c+"次的数据为："+dataGridView.Rows[d].Cells[3].Value.ToString());
                                    }
                                    //计数器加上数组的长度，用于从第N列重新开始赋值
                                    counter = counter + lujingshuzu.Length;
                                }
                                #endregion

                            }
                            #endregion
                        }
                        
                    }
                    else
                    {
                        #region 进行路径拆分
                        //循环将路径那一列并进行拆分
                        for (int c = 0; c < Int32.Parse(counters[g]) - Int32.Parse(counters[g-1]); c++)
                        {
                            //MessageBox.Show(array[c,3]);

                            /**
                             * 将路径的那一列，根据“；”进行拆分,    StringSplitOptions.RemoveEmptyEntries的作用为去除空值
                             * 因为999和并签的部分环节对应的可提交路径为空，所以此处会把999和并签的部分环节给去掉,
                             * */
                            lujingshuzu = array[Int32.Parse(counters[g - 1])+1+c, 5].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                            //将路径的那一列，根据“；”进行拆分。   因为999的可提交路径为空，所以不能在此去除空值，否则会导致999无法被写入，无法被转换
                            //lujingshuzu = array[c, 3].Split(new char[] { ';' });




                            #region  此段代码作用是二维数组的数据做一定处理后，赋给datatable

                            /**
                             * 根据数据内有没有值，也就是数组的长度不同，有不同的逻辑控制
                            **/


                            //如果数组内没有值，也就是数组的长度为0
                            if (lujingshuzu.Length == 0)
                            {
                                //定义一个dataTable的行
                                DataRow dr = dt.NewRow();
                                //为dataTable增加一行
                                dt.Rows.Add(dr);

                                //给datatable的第counter行的6列赋值
                                dt.Rows[counter][0] = counter + 1;//序号
                                dt.Rows[counter][1] = array[Int32.Parse(counters[g - 1]) + c + 1, 1];//流程标识
                                dt.Rows[counter][2] = array[Int32.Parse(counters[g - 1]) + c + 1, 2];//流程名称
                                dt.Rows[counter][3] = array[Int32.Parse(counters[g - 1]) + c + 1, 3];//环节ID
                                dt.Rows[counter][4] = array[Int32.Parse(counters[g - 1]) + c + 1, 4];//环节名称
                                dt.Rows[counter][5] = "";//可提交路径

                                //计数器加上数组的长度，用于从第N行重新开始赋值
                                counter = counter + 1;

                            }
                            else//数组内有值，也就是数组的长度大于0
                            {
                                //循环将数组的值赋给datatable
                                for (int d = 0; d < lujingshuzu.Length; d++)
                                {
                                    //定义一个dataTable的行
                                    DataRow dr = dt.NewRow();
                                    //为dataTable增加一行
                                    dt.Rows.Add(dr);


                                    //dataGridView.Rows.Add();
                                    //定义一个临时计数器
                                    int counterls = counter + d;
                                    //临时测试，展示C的值
                                    //MessageBox.Show("C的值为" + c );


                                    //临时测试，展示lujingshuzu的值
                                    //for (int e = 0; e < lujingshuzu.Length;e++ )
                                    //{
                                    //    MessageBox.Show(lujingshuzu[e]);
                                    //}



                                    //判断根据；拆分后的值是否为空，不为空则赋值，如果为空则跳出循环
                                    if (lujingshuzu[d] != null && lujingshuzu[d] != "")
                                    {
                                        //dataGridView.Rows[d].Cells[1].Value = array[c, 1];
                                        //dataGridView.Rows[d].Cells[2].Value = array[c, 2];   
                                        //dataGridView.Rows[d].Cells[3].Value = lujingshuzu[d];

                                        ////给datatable的第N行的6列赋值
                                        dt.Rows[counterls][0] = counterls + 1;//序号
                                        dt.Rows[counterls][1] = array[Int32.Parse(counters[g-1])+c+1, 1];//流程标识
                                        dt.Rows[counterls][2] = array[Int32.Parse(counters[g - 1]) + c + 1, 2];//流程名称
                                        dt.Rows[counterls][3] = array[Int32.Parse(counters[g - 1]) + c + 1, 3];//环节ID
                                        dt.Rows[counterls][4] = array[Int32.Parse(counters[g - 1]) + c + 1, 4];//环节名称
                                        dt.Rows[counterls][5] = lujingshuzu[d];//可提交路径

                                        //临时测试，展示拆分后的数据
                                        //MessageBox.Show("第C" + c + "次的数据为：" + dataGridView.Rows[d].Cells[3].Value.ToString());
                                    }
                                    else
                                    {
                                        continue;
                                    }


                                    //临时测试，展示拆分后的数据
                                    //MessageBox.Show("第"+c+"次的数据为："+dataGridView.Rows[d].Cells[3].Value.ToString());
                                }
                                //计数器加上数组的长度，用于从第N列重新开始赋值
                                counter = counter + lujingshuzu.Length;
                            }
                            #endregion

                        }
                        #endregion
                    }



                }
                #endregion





                #region  此段代码已废止，此段代码只支持非并签的数据
                ////循环获取所有的可提交路径
                //for (i = 0; i < dataGridView.Rows.Count; i++)
                //{
                //    //获取可提交路径，并将值赋给一个String字段，因为可提交路径为第三列，也就是下标为2的那一列
                //    lujing = dataGridView.Rows[i].Cells[3].Value.ToString();

                //    //循环
                //    for (j = 0; j < dataGridView.Rows.Count; j++)
                //    {
                //        //将之前获取到的可提交路径的 00几，替换为文字说明，此处有BUG，暂不支持并签。
                //        lujing = lujing.Replace(dataGridView.Rows[j].Cells[1].Value.ToString(), dataGridView.Rows[j].Cells[2].Value.ToString());

                //    }
                //    //去掉路径中的左半边大括号
                //    lujing = lujing.Replace("{","");
                //    //去掉路径中的左半边大括号
                //    lujing = lujing.Replace("}", "");

                //    //将修改过的可提交路径重新写入到dataGridView中
                //    dataGridView.Rows[i].Cells[2].Value = lujing;
                //}
                #endregion


                #region  此段代码已废止，此段代码只支持并签的数据,用的是dataGridView，

                ////循环获取所有的可提交路径
                //for (i = 0; i < dataGridView.Rows.Count; i++)
                //{   

                //    //获取可提交路径，并将值赋给一个String字段，因为可提交路径为第三列，也就是下标为2的那一列
                //    lujing = dataGridView.Rows[i].Cells[3].Value.ToString();

                //    //将左半边大括号替换为 右半边的大括号，以便进行数据拆分
                //    lujing = lujing.Replace("{", "}");

                //    //string[] s = str.Split(new char[] { '#' });参考代码
                //    //将路径根据右版本的大括号拆分成一个数据，用于数据比对
                //    lujingshuzu = lujing.Split(new char[] { '}' });

                //    //因为会出现重复数据，不知道怎么处理的关系，所以在此将lujing的值 置空
                //    lujing = null;

                //    //以下代码的作用为将可提交路径内的00几替换为文字
                //    //根据数组的长度循环
                //    for (k = 0; k < lujingshuzu.Length;k++ )
                //    {
                //        //根据导入excel的行数循环
                //        for (j = 0; j < dataGridView.Rows.Count; j++)
                //        {
                //            //循环判断数组内的00几和可提交路径列的00及是否匹配，如果匹配则替换为文字说明
                //            if (lujingshuzu[k].Equals(dataGridView.Rows[j].Cells[1].Value.ToString()))
                //            {
                //                lujingshuzu[k] = dataGridView.Rows[j].Cells[2].Value.ToString();
                //            }

                //        }

                //        lujing = lujing + lujingshuzu[k];
                //    }
                //    //删除#A#
                //    lujing = lujing.Replace("#A#", "");
                //    //删除#N#
                //    lujing = lujing.Replace("#N#", "");
                //    //删除#T#
                //    lujing = lujing.Replace("#T#", "");
                //    //将修改过的可提交路径重新写入到dataGridView中
                //    dataGridView.Rows[i].Cells[3].Value = lujing;
                //}
                #endregion


                //return dt;


                #region  此段代码支持并签的数据,用的是datatable，作用是将【00X】替换为文字说明



                //循环获取所有的可提交路径
                for (i = 0; i < dt.Rows.Count; i++)
                {

                    //获取可提交路径，并将值赋给一个String字段，因为可提交路径为第6列，也就是下标为5的那一列
                    lujing = dt.Rows[i][5].ToString();

                    //将左半边大括号替换为 右半边的大括号，以便进行数据拆分
                    lujing = lujing.Replace("{", "}");

                    //string[] s = str.Split(new char[] { '#' });参考代码
                    //将路径根据右版本的大括号拆分成一个数据，用于数据比对
                    lujingshuzu = lujing.Split(new char[] { '}' });

                    //因为会出现重复数据，不知道怎么处理的关系，所以在此将lujing的值 置空
                    lujing = null;

                    //以下代码的作用为将可提交路径内的00几替换为文字

                    //临时测试，展示拆分后的数据
                    //MessageBox.Show("datatable的行数为：" + dt.Rows.Count.ToString());

                    //根据数组的长度循环
                    for (k = 0; k < lujingshuzu.Length; k++)
                    {
                        //根据导入excel的行数循环
                        for (j = 0; j < dt.Rows.Count; j++)
                        {
                            //循环判断数组内的00几和可提交路径列的00及是否匹配，如果匹配则替换为文字说明
                            if (lujingshuzu[k].Equals(dt.Rows[j][3].ToString()))
                            {
                                lujingshuzu[k] = dt.Rows[j][4].ToString();
                            }

                        }

                        lujing = lujing + lujingshuzu[k];
                    }
                    //删除#A#
                    //lujing = lujing.Replace("#A#", "");
                    //删除#N#
                    //lujing = lujing.Replace("#N#", "");
                    //删除#T#
                    //lujing = lujing.Replace("#T#", "");
                    //删除WORKFLOWSECRSELE
                    //lujing = lujing.Replace("WORKFLOWSECRSELE", "");

                    //将修改过的可提交路径重新写入到dataTable中
                    dt.Rows[i][5] = lujing;
                }
                #endregion

                //返回一个datatable
                return dt;



            }
            catch (Exception ex)
            {
                //抛出异常
                MessageBox.Show("更新Excel失败!失败原因：" + ex.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }


        /// <summary>
        /// 加载Excel
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>返回一个DataSet</returns>
        public static DataSet LoadDataFromExcel(string filePath)
        {
            try
            {

                //连接字符串
                string connStr;

                //根据文件的路径，获取excel的格式，包含.
                string fileType = System.IO.Path.GetExtension(filePath);

                //03格式和07格式的excel连接字符串是不一样的，所以需要进行判断
                if (string.IsNullOrEmpty(fileType))
                {
                    MessageBox.Show("绑定数据异常");
                    //非空判断，如果为空抛出null
                    return null;
                }
                else if (fileType == ".xls")//03格式
                {
                    //03格式的Excel
                    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
                }

                else
                {
                    //07格式的Excel
                    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
                }

                //string sql = "Select * FROM [27$]";//直接指定sheet页的名字就不报错
                ////创建一个DataTable用于存放excel中的数据
                //DataTable dataTable = null;

                //用于连接数据库
                OleDbConnection conn = null;

                //创建一个DataAdapter 用于填充DataSet
                OleDbDataAdapter dataAdapter = null;

                //创建一个DataSet用于填充
                DataSet dataSet = null;

                try
                {
                    // 初始化连接 
                    conn = new OleDbConnection(connStr);

                    //打开数据库
                    conn.Open();

                    //获取excel页签的个数
                    //string[] names = GetExcelSheetNames(conn);

                    //临时测试用，查看sheet的名字都是什么
                    //for (int i = 0; i < names.Length;i++ )
                    //{
                    //    MessageBox.Show(names[i]);
                    //}

                    //不知道为什么动态获取 names[0]的数据会有问题，所以在此定义一个值，用于存储names[0]的值
                    //String sheetName = names[0];

                    //MessageBox.Show(sheetName);

                    /**
                     * string sql = "Select * FROM [{0}]";
                     * string sql = "Select * FROM [SQL Results$]"
                     * 将excel当做一张表查询  【{0}】标表示Excel页签的位置
                     * excel的格式为  SQL Results$
                     **/

                    //拼接sql，根据excel第1个sheet的名字，拼接sql
                    //string sql = "Select * FROM [" + sheetName + "]";
                    string sql = "Select * FROM [SQL Results$]";

                    //实例化一个dataAdapter
                    dataAdapter = new OleDbDataAdapter(sql, conn);
                    //实例化一个dataSet
                    dataSet = new DataSet();

                    //填充
                    //dataAdapter.Fill(dataSet, sheetName);
                    dataAdapter.Fill(dataSet, "SQL Results$");

                    //关闭连接
                    conn.Close();


                    //MessageBox.Show("数据绑定成功");
                    //返回存储excel信息的dataSet
                    return dataSet;


                }
                catch (Exception ex)
                {
                    //抛出异常
                    MessageBox.Show("数据绑定Excel失败!失败原因：" + ex.Message, "提示信息",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }
                finally
                {
                    // 判断conn的状态是否为Open        
                    if (conn.State == ConnectionState.Open)
                    {
                        //关闭数据库
                        conn.Close();
                        //释放dataAdapter使用的资源
                        dataAdapter.Dispose();
                        //释放conn使用的资源
                        conn.Dispose();
                    }
                }
            }

            catch (Exception ex)
            {
                //抛出异常
                MessageBox.Show("数据绑定Excel失败!失败原因：" + ex.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }



        /// <summary>
        /// 直接将DataTable输出为excel
        /// </summary>
        /// <param name="excelTable">参数DataTable</param>
        /// <param name="filePath">参数，excel存放的路径</param>
        /// <returns>返回值，返回一个成功或失败（true、false）</returns>
        public static bool SaveDataTableToExcel(System.Data.DataTable excelTable, string filePath)
        {

            Object missing = Missing.Value;
            //定义计数器1,定义起始位置
            int counterOns = 0;
            //定义计数器2，定义结束位置
            int conterTwo = 1;


            try
            {
                //判断是否有数据，如果没有数据弹出提示 
                if (excelTable.Rows.Count == 0)
                {
                    MessageBox.Show("当前没有数据！");
                    return false;
                }

                //需要在引用中的.NET中添加【Microsoft.Office.Interop.Excel】,并将【Microsoft.Office.Interop.Excel】的嵌入操作类型改为False
                Microsoft.Office.Interop.Excel.Application app =
                    new Microsoft.Office.Interop.Excel.ApplicationClass();

                //Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8;

                // 判断电脑是否装了excel，如果没装，弹出如下提示
                if (app == null)
                {
                    MessageBox.Show("无法创建Excel对象，可能未安装Excel");
                    return false;
                }
                //获取当前时间，用于杀死进程
                beforeTime = DateTime.Now;
                //false为让后台执行设置为不可见，为true的话会看到打开一个Excel，然后数据在往里写
                app.Visible = false;

                //excel文档
                Workbook wBook = app.Workbooks.Add(true);


                //excel中的一个页签
                Worksheet wSheet;

                //用于设置excel的格式
                //Range range;


                //定义一个集合
                List<String> counters = new List<string>();
                //获取流程的数量
                counters = cordysNumber(excelTable);

                if (counters==null)
                {
                    MessageBox.Show("未找到需要转换的流程");
                    return false;
                }

                


                //根据数组的个数判断流程的个数
                for (int i = 1; i < counters.Count(); i++)
                {
                    wSheet = (Microsoft.Office.Interop.Excel.Worksheet)wBook.Worksheets.get_Item(i);
                    wSheet.Copy(missing, wBook.Worksheets[i]);
                }

                for (int j = 0; j < counters.Count();j++ )
                {

                    //第一次的逻辑和其他的不同，所以进行了逻辑判断
                    if (j == 0)
                    {
                        //获取要写入数据的WorkSheet对象，并重命名
                        wSheet = (Worksheet)wBook.Worksheets.get_Item(j + 1);


                        #region  合并指定列的相同的单元格

                        //重置计数器1,定义起始位置
                        counterOns = 0;
                        //重置计数器2，定义结束位置
                        conterTwo = 1;

                        //合并指定列的相同的单元格

                        for (int i = 0; i < Int32.Parse(counters[j]); i++)
                        {
                            //MessageBox.Show("第" + counterOns  + "行的值为：" + excelTable.Rows[counterOns][4].ToString());
                            //MessageBox.Show("第" + conterTwo  + "行的值为：" + excelTable.Rows[conterTwo][4].ToString());

                            //判断“环节名称”y以及“环节ID”是否相同
                            if (excelTable.Rows[counterOns][4].ToString().Equals(excelTable.Rows[conterTwo][4].ToString())
                                && excelTable.Rows[counterOns][3].ToString().Equals(excelTable.Rows[conterTwo][3].ToString()))
                            {

                                //Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                                //application.DisplayAlerts = false;
                                //环节ID列单合并元格，
                                wSheet.Cells.get_Range(wSheet.Cells[counterOns + 2, 4], wSheet.Cells[conterTwo + 2, 4]).Merge();
                                //环节名称列合并单元格，
                                wSheet.Cells.get_Range(wSheet.Cells[counterOns + 2, 5], wSheet.Cells[conterTwo + 2, 5]).Merge();

                                conterTwo = conterTwo + 1;
                            }
                            else
                            {
                                counterOns = conterTwo;
                                conterTwo = conterTwo + 1;
                            }

                        }


                        //流程ID列合并单元格，
                        wSheet.Cells.get_Range(wSheet.Cells[2, 2], wSheet.Cells[Int32.Parse(counters[j])+1, 2]).Merge();
                        //流程名称列合并单元格，
                        wSheet.Cells.get_Range(wSheet.Cells[2, 3], wSheet.Cells[Int32.Parse(counters[j])+1, 3]).Merge();

                        #endregion


                        //判断counters[j]是否大于0（也就是判断DataTable内是否有数据），如果有数据则输出到excel
                        if (Int32.Parse(counters[j]) > 0)
                        {
                            #region  此段代码作用是输出excel的头信息，也就是第一行,循环输出，已废止
                            ////获取datatable的列的长度
                            //int size = excelTable.Columns.Count;
                            ////循环写入列名
                            //for (int i = 1; i < size; i++)
                            //{
                            //    wSheet.Cells[1, i + 1] = excelTable.Columns[i].ColumnName;
                            //}
                            ////定义第一行第一列的名字为序号
                            //wSheet.Cells[1, 1] = "序号";
                            #endregion

                            #region  此段代码作用是输出excel的头信息，也就是第一行，指定列名

                            //定义第一行第N列的名字
                            wSheet.Cells[1, 1] = "序号";
                            wSheet.Cells[1, 2] = "流程ID";
                            wSheet.Cells[1, 3] = "流程名称";
                            wSheet.Cells[1, 4] = "环节ID";
                            wSheet.Cells[1, 5] = "环节名称";
                            wSheet.Cells[1, 6] = "可提交路径";

                            #endregion

                            #region 此段代码作用是输出excel除头头信息以外的数据
                            //用于存储的行数
                            int row = Int32.Parse(counters[j])+1;

                            //获取的列数
                            int col = excelTable.Columns.Count;


                            #region 此段代码作用是循环输出excel，不指定列，循环输出  已废止
                            ////此方式不用管有几列，全部循环输出，但是以为数字的关系，此方式不太合适
                            ////循环将datatable的数据写入到excel，不包括列名的信息
                            //for (int i = 0; i < row; i++)
                            //{



                            //    for (int j = 0; j < col; j++)
                            //    {
                            //        //因为输出的时候不需要输出999，所以判断是否为999，如果是999则跳出循环
                            //        if (excelTable.Rows[i][3].ToString() == "999" || excelTable.Rows[i][3].ToString().Equals("999"))
                            //        {
                            //            break;
                            //        }
                            //        else
                            //        {


                            //            //获取datatable的某个单元格的值
                            //            string str = excelTable.Rows[i][j].ToString();
                            //            //将值写入到excel中，从第二行，第一列开始写入，excel开始为1而不是0，datatable有区别
                            //            wSheet.Cells[i + 2, j + 1] = str;
                            //        }

                            //    }

                            //}
                            #endregion



                            #region 此段代码作用是循环输出excel，指定列，循环输出。
                            //循环将datatable的数据写入到excel，不包括列名的信息
                            for (int i = 0; i < row; i++)
                            {
                                //因为输出的时候不需要输出999，所以判断是否为999，如果是999则跳出循环
                                if (excelTable.Rows[i][3].ToString() == "999" || excelTable.Rows[i][3].ToString().Equals("999"))
                                {
                                    break;
                                }
                                else
                                {
                                    //将值写入到excel中，从第二行，第一列开始写入，excel开始为1而不是0，datatable有区别
                                    wSheet.Cells[i + 2, 1] = "'" + excelTable.Rows[i][0].ToString();//序号，前面加【'】是为了强转为文本格式
                                    wSheet.Cells[i + 2, 2] = "'" + excelTable.Rows[i][1].ToString();//流程ID，前面加【'】是为了强转为文本格式
                                    wSheet.Cells[i + 2, 3] = excelTable.Rows[i][2].ToString();
                                    wSheet.Cells[i + 2, 4] = "'" + excelTable.Rows[i][3].ToString();//环节ID，前面加【'】是为了强转为文本格式
                                    wSheet.Cells[i + 2, 5] = excelTable.Rows[i][4].ToString();
                                    wSheet.Cells[i + 2, 6] = excelTable.Rows[i][5].ToString();

                                }
                            }
                            #endregion

                            #endregion
                        }

                        //重命名sheet页的名字流程名字+流程ID的组合
                        wSheet.Name = excelTable.Rows[Int32.Parse(counters[j])][2].ToString()
                            + "(" 
                            + excelTable.Rows[Int32.Parse(counters[j])][1].ToString()
                            +")";
                    }
                    else 
                    {
                        //获取要写入数据的WorkSheet对象，并重命名
                        wSheet = (Worksheet)wBook.Worksheets.get_Item(j + 1);


                        #region  合并指定列的相同的单元格

                        //重置计数器1,定义数据起始位置
                        counterOns = Int32.Parse(counters[j - 1]) + 1;
                        //重置计数器2，定义数据结束位置
                        conterTwo = Int32.Parse(counters[j - 1]) + 2;

                        //重置计数器3,定义excel起始位置
                        int counthree = 0;
                        //重置计数器4，定义excel结束位置
                        int conterfour = 1;

                        //合并指定列的相同的单元格

                        for (int i = 0; i < Int32.Parse(counters[j]) - Int32.Parse(counters[j - 1])-1; i++)
                        {
                            //MessageBox.Show("第" + counterOns  + "行的值为：" + excelTable.Rows[counterOns][4].ToString());
                            //MessageBox.Show("第" + conterTwo  + "行的值为：" + excelTable.Rows[conterTwo][4].ToString());

                            //判断“环节名称”以及“环节ID”是否相同
                            if (excelTable.Rows[counterOns][4].ToString().Equals(excelTable.Rows[conterTwo][4].ToString())
                                && excelTable.Rows[counterOns][3].ToString().Equals(excelTable.Rows[conterTwo][3].ToString()))
                            {

                                //Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                                //application.DisplayAlerts = false;
                                //环节ID列单合并元格，
                                wSheet.Cells.get_Range(wSheet.Cells[counthree + 2, 4], wSheet.Cells[conterfour + 2, 4]).Merge();
                                //环节名称列合并单元格，
                                wSheet.Cells.get_Range(wSheet.Cells[counthree + 2, 5], wSheet.Cells[conterfour + 2, 5]).Merge();

                                conterTwo = conterTwo + 1;
                                conterfour = conterfour + 1;
                            }
                            else
                            {
                                counterOns = conterTwo;
                                conterTwo = conterTwo + 1;

                                counthree = conterfour;
                                conterfour=conterfour + 1;
                            }

                        }


                        //流程ID列合并单元格，
                        wSheet.Cells.get_Range(wSheet.Cells[2, 2], wSheet.Cells[Int32.Parse(counters[j]) - Int32.Parse(counters[j - 1]) , 2]).Merge();
                        //流程名称列合并单元格，
                        wSheet.Cells.get_Range(wSheet.Cells[2, 3], wSheet.Cells[Int32.Parse(counters[j]) - Int32.Parse(counters[j - 1]) , 3]).Merge();

                        #endregion


                        //判断counters[j]是否大于0（也就是判断DataTable内是否有数据），如果有数据则输出到excel
                        if (Int32.Parse(counters[j]) > 0)
                        {
                            #region  此段代码作用是输出excel的头信息，也就是第一行,循环输出，已废止
                            ////获取datatable的列的长度
                            //int size = excelTable.Columns.Count;
                            ////循环写入列名
                            //for (int i = 1; i < size; i++)
                            //{
                            //    wSheet.Cells[1, i + 1] = excelTable.Columns[i].ColumnName;
                            //}
                            ////定义第一行第一列的名字为序号
                            //wSheet.Cells[1, 1] = "序号";
                            #endregion

                            #region  此段代码作用是输出excel的头信息，也就是第一行，指定列名

                            //定义第一行第N列的名字
                            wSheet.Cells[1, 1] = "序号";
                            wSheet.Cells[1, 2] = "流程ID";
                            wSheet.Cells[1, 3] = "流程名称";
                            wSheet.Cells[1, 4] = "环节ID";
                            wSheet.Cells[1, 5] = "环节名称";
                            wSheet.Cells[1, 6] = "可提交路径";

                            #endregion

                            #region 此段代码作用是输出excel除头头信息以外的数据
                            //用于存储的行数
                            int row = Int32.Parse(counters[j])-Int32.Parse(counters[j-1]);

                            //获取的列数
                            int col = excelTable.Columns.Count;


                            #region 此段代码作用是循环输出excel，不指定列，循环输出  已废止
                            ////此方式不用管有几列，全部循环输出，但是以为数字的关系，此方式不太合适
                            ////循环将datatable的数据写入到excel，不包括列名的信息
                            //for (int i = 0; i < row; i++)
                            //{



                            //    for (int j = 0; j < col; j++)
                            //    {
                            //        //因为输出的时候不需要输出999，所以判断是否为999，如果是999则跳出循环
                            //        if (excelTable.Rows[i][3].ToString() == "999" || excelTable.Rows[i][3].ToString().Equals("999"))
                            //        {
                            //            break;
                            //        }
                            //        else
                            //        {


                            //            //获取datatable的某个单元格的值
                            //            string str = excelTable.Rows[i][j].ToString();
                            //            //将值写入到excel中，从第二行，第一列开始写入，excel开始为1而不是0，datatable有区别
                            //            wSheet.Cells[i + 2, j + 1] = str;
                            //        }

                            //    }

                            //}
                            #endregion



                            #region 此段代码作用是循环输出excel，指定列，循环输出。
                            //循环将datatable的数据写入到excel，不包括列名的信息
                            for (int i = 0; i < row; i++)
                            {
                                //因为输出的时候不需要输出999，所以判断是否为999，如果是999则跳出循环
                                if (excelTable.Rows[Int32.Parse(counters[j - 1]) + 1 + i][3].ToString() == "999" || excelTable.Rows[Int32.Parse(counters[j - 1]) + 1 + i][3].ToString().Equals("999"))
                                {
                                    break;
                                }
                                else
                                {
                                    //将值写入到excel中，从第二行，第一列开始写入，excel开始为1而不是0，datatable有区别
                                    wSheet.Cells[i + 2, 1] = "'" + (i+1).ToString();//序号，前面加【'】是为了强转为文本格式
                                    wSheet.Cells[i + 2, 2] = "'" + excelTable.Rows[Int32.Parse(counters[j - 1]) + 1 + i][1].ToString();//流程ID，前面加【'】是为了强转为文本格式
                                    wSheet.Cells[i + 2, 3] = excelTable.Rows[Int32.Parse(counters[j - 1]) + 1 + i][2].ToString();
                                    wSheet.Cells[i + 2, 4] = "'" + excelTable.Rows[Int32.Parse(counters[j - 1]) + 1 + i][3].ToString();//环节ID，前面加【'】是为了强转为文本格式
                                    wSheet.Cells[i + 2, 5] = excelTable.Rows[Int32.Parse(counters[j - 1]) + 1 + i][4].ToString();
                                    wSheet.Cells[i + 2, 6] = excelTable.Rows[Int32.Parse(counters[j - 1]) + 1 + i][5].ToString();

                                }
                            }
                            #endregion

                            #endregion
                        }

                        //重命名sheet页的名字流程名字+流程ID的组合
                        wSheet.Name = excelTable.Rows[Int32.Parse(counters[j])][2].ToString()
                            + "("
                            + excelTable.Rows[Int32.Parse(counters[j])][1].ToString()
                            + ")";
                    }
                }



       


               
                //wSheet.get_Range(wSheet.Cells[2,3],wSheet.Cells[6,3]).Merge();

                //设置禁止弹出保存和覆盖的询问提示框 
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;

                //保存工作簿(也就是常说的页签) 
                wBook.Save();

                //保存excel文件 
                app.Save(filePath);


                //获取当前时间用于杀死进程
                afterTime = DateTime.Now;

                //不知道这句话是什么意思
                //app.SaveWorkspace(filePath);

                //貌似是清空什么东西
                app.Quit();
                app = null;

                MessageBox.Show("导出成功");
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                System.GC.Collect();
                KillExcelProcess();
                return false;

            }
            finally
            {
                //KillProcess("EXCEL");//杀死进程EXCEL 
                //int generation = System.GC.GetGeneration(excel);
                //excelApp = null;
                System.GC.Collect();
                KillExcelProcess();

                //MessageBox.Show(ex.Message, "错误提示"); 


            }


        }


        /// <summary>
        /// 控制DataGridView为不可排序的状态，此方法用在DataGridView的数据加载完成之后
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        public static void ForbidSortColumn(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }




        /// <summary>
        /// 查询excel的Sheet页的名字
        /// </summary>
        /// <param name="con"></param>
        /// <returns></returns>
        public static string[] GetExcelSheetNames(OleDbConnection con)
        {
            try
            {
                /**
                 * 检索Excel的架构信息
                 * 百度直接复制的，不懂为什么这样用，不懂这个方法的原理
                 * 只知道大概意思为将excel的架构放到一个datatable中
                 * 
                 **/
                System.Data.DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new[] { null, null, null, "Table" });

                //简历一个数组，数组的长度为sheet页的个数
                String[] sheet = new String[dt.Rows.Count];

                //循环将Sheet页的名字赋给数组
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet[i] = dt.Rows[i]["TABLE_NAME"].ToString();
                }
                return sheet;
            }
            catch
            {
                return null;
            }
        }



        //private static String[] DeleteNull(String[] lujingshuzu)
        //{
        //    String[] lujingshuzulinshi =  new String[]{};

        //    for (int i = 0; i < lujingshuzu.Length;i++ )
        //    {
        //        if(lujingshuzu[i]!=null&&lujingshuzu[i]!="")
        //        {

        //        }
        //        else
        //        {
        //            continue;
        //        }
        //    }
        //}


        /// <summary>
        /// 判断有多少个流程
        /// </summary>
        /// <param name="excelTable">输入参数为Datatable</param>
        /// <returns>返回一个List<String></returns>
        private static List<String> cordysNumber(System.Data.DataTable excelTable)
        {
           if(excelTable.Rows.Count>0)
           {
               //定义一个计数器集合
               List<String> counters = new List<string>();
               //String[] counters = new String[] { };

               //定义计数器1,定义起始位置
               int counterOns = 0;
               //定义计数器2，定义结束位置
               int conterTwo = 1;
               //定义计数器3，当只有一个流程是用来定义流程个数
               int conterThree = 1;

               //MessageBox.Show(dataGridView.Rows.Count.ToString());

               //循环判断有多少个流程
               for (int g = 0; g < excelTable.Rows.Count - 1; g++)
               {
                   //判断“流程ID”那一列的值是否相同
                   if (excelTable.Rows[counterOns][1].ToString().Equals(excelTable.Rows[conterTwo][1].ToString()))
                   {
                       conterTwo = conterTwo + 1;
                   }
                   else
                   {
                       counterOns = conterTwo;
                       conterTwo = conterTwo + 1;
                       conterThree = conterThree + 1;
                       //将同一个流程最后一行的下标赋给数组
                       counters.Add((counterOns - 1).ToString());
                   }
               }

               //如果只有一个流程那么上方的给数组赋值则不生效，此段代码是当只有一个流程时，给数组赋值
               if (conterThree == 1)
               {
                   //MessageBox.Show((conterTwo).ToString());
                   counters.Add((conterTwo - 1).ToString());
               }
               else//否则将最后一个赋给数组
               {
                   counters.Add((conterTwo - 1).ToString());
               }

               return counters;
           }
           else
           {
               return null;
           }
        }




        /// <summary>
        /// 判断有多少个流程
        /// </summary>
        /// <param name="excelTable">输入参数为DataGridView</param>
        /// <returns>返回一个List<String></returns>
        private static List<String> cordysNumber(DataGridView dataGridView)
        {
            if (dataGridView.Rows.Count > 0)
            {
                #region  判断相同的流程

                //定义一个计数器集合
                List<String> counters = new List<string>();
                //String[] counters = new String[] { };

                //定义计数器1,定义起始位置
                int counterOns = 0;
                //定义计数器2，定义结束位置
                int conterTwo = 1;
                //定义计数器3，当只有一个流程是用来定义流程个数
                int conterThree = 1;

                //MessageBox.Show(dataGridView.Rows.Count.ToString());

                //循环判断有多少个流程
                for (int g = 0; g < dataGridView.Rows.Count - 1; g++)
                {
                    //判断“流程ID”那一列的值是否相同
                    if (dataGridView.Rows[counterOns].Cells[1].Value.ToString().Equals(dataGridView.Rows[conterTwo].Cells[1].Value.ToString()))
                    {
                        conterTwo = conterTwo + 1;
                    }
                    else
                    {
                        counterOns = conterTwo;
                        conterTwo = conterTwo + 1;
                        conterThree = conterThree + 1;
                        //将同一个流程最后一行的下标赋给数组
                        counters.Add((counterOns - 1).ToString());
                    }
                }

                //如果只有一个流程那么上方的给数组赋值则不生效，此段代码是当只有一个流程时，给数组赋值
                if (conterThree == 1)
                {
                    //MessageBox.Show((conterTwo).ToString());
                    counters.Add((conterTwo - 1).ToString());
                }
                else//否则将最后一个赋给数组
                {
                    counters.Add((conterTwo - 1).ToString());
                }
                #endregion

                return counters;
            }
            else
            {
                return null;
            }
        }


        /// <summary>
        /// 根据用户选择的文件获取文件的路径。
        /// 如果获取文件路径成功则返回路径，如果失败则返回mull
        /// /// <param name="openFileDialog">OpenFileDialog</param>
        /// </summary>
        /// <returns>如果获取文件路径成功则返回路径，如果失败则返回mull</returns>
        public static String OpenFileDialog(OpenFileDialog openFileDialog)
        {
            //展示提示语言
            openFileDialog.Title = "请选择Excel，目前仅支持03和07格式的Excel，不支持WPS";
            //规定用户可以选择文件的类型，用后缀名控制，只支持【.xlsx】和【.xls】
            openFileDialog.Filter = "Excel文件(*.xlsx;*.xls)|*.xlsx;*.xls";
            //控制用户是否可以选择多个文件，ture为可以选择多个，false为不可以选择多个；
            openFileDialog.Multiselect = false;

            openFileDialog.ValidateNames = true;     //文件有效性验证ValidateNames，验证用户输入是否是一个有效的Windows文件名
            openFileDialog.CheckFileExists = true;  //验证路径有效性
            openFileDialog.CheckPathExists = true; //验证文件有效性


            //判断用户是否点击了Ok按钮
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取文件的的名字
                String FileName = openFileDialog.FileNames[0];
                //判断文件名是否为空
                if (FileName == "" || FileName.Equals(null))
                {
                    return null;

                }
                else
                {
                    return FileName;
                }
            }
            else
            {
                return null;
            }
        }

        
        /// <summary>
        /// 根据用户的选择获取保存文件的路径
        /// </summary>
        /// <param name="saveFileDialog">SaveFileDialog</param>
        /// <returns></returns>
        public static String SaveFileDialog(SaveFileDialog saveFileDialog)
        {

            //saveFileDialog.Filter = "导出Excel (*.xls)|*.xls|导出Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Filter = "导出Excel (*.xls)|*.xls";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "导出文件保存路径，目前仅支持03和07格式的Excel，不支持WPS";

            




            //判断用户是否点击了Ok按钮
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取文件的的名字
                String FileName = saveFileDialog.FileNames[0];
                //判断文件名是否为空
                if (FileName == "" || FileName.Equals(null))
                {
                    return null;

                }
                else
                {
                    return saveFileDialog.FileName;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 结束Excel进程
        /// </summary>
        public static void KillExcelProcess()
        {
            Process[] myProcesses;
            DateTime startTime;
            myProcesses = Process.GetProcessesByName("Excel");

            //得不到Excel进程ID，暂时只能判断进程启动时间
            foreach (Process myProcess in myProcesses)
            {
                startTime = myProcess.StartTime;

                if (startTime > beforeTime && startTime < afterTime)
                {
                    myProcess.Kill();
                }
            }
        }



        #region 废掉的 没有通用性
        ///// <summary>
        ///// 合并单元格
        ///// </summary>
        ///// <param name="j">数组的位置</param>
        ///// <param name="counterOns"></param>
        ///// <param name="conterTwo"></param>
        ///// <param name="counters"></param>
        ///// <param name="excelTable"></param>
        ///// <param name="wSheet"></param>
        //private static void merge(int j, int counterOns, int conterTwo, List<String> counters, System.Data.DataTable excelTable, Worksheet wSheet)
        //{

        //    //计数器1,定义起始位置
        //    counterOns = 0;
        //    //计数器2，定义结束位置
        //    conterTwo = 1;

        //    //合并指定列的相同的单元格

        //    for (int i = 0; i < Int32.Parse(counters[j]); i++)
        //    {
        //        //MessageBox.Show("第" + counterOns  + "行的值为：" + excelTable.Rows[counterOns][4].ToString());
        //        //MessageBox.Show("第" + conterTwo  + "行的值为：" + excelTable.Rows[conterTwo][4].ToString());

        //        //判断“环节名称”y以及“环节ID”是否相同
        //        if (excelTable.Rows[counterOns][4].ToString().Equals(excelTable.Rows[conterTwo][4].ToString())
        //            && excelTable.Rows[counterOns][3].ToString().Equals(excelTable.Rows[conterTwo][3].ToString()))
        //        {

        //            //Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
        //            //application.DisplayAlerts = false;
        //            //环节ID列单合并元格，
        //            wSheet.Cells.get_Range(wSheet.Cells[counterOns + 2, 4], wSheet.Cells[conterTwo + 2, 4]).Merge();
        //            //环节名称列合并单元格，
        //            wSheet.Cells.get_Range(wSheet.Cells[counterOns + 2, 5], wSheet.Cells[conterTwo + 2, 5]).Merge();

        //            conterTwo = conterTwo + 1;
        //        }
        //        else
        //        {
        //            counterOns = conterTwo;
        //            conterTwo = conterTwo + 1;
        //        }

        //    }

        //    //流程ID列合并单元格，
        //    wSheet.Cells.get_Range(wSheet.Cells[2, 2], wSheet.Cells[Int32.Parse(counters[j]) + 1, 2]).Merge();
        //    //流程名称列合并单元格，
        //    wSheet.Cells.get_Range(wSheet.Cells[2, 3], wSheet.Cells[Int32.Parse(counters[j]) + 1, 3]).Merge();

        //}
        #endregion

    }
}