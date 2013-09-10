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

namespace Excel
{
    class LinShi
    {
        private void ToExcelSheet(DataSet ds,string sheetName)
        {
            int testnum = ds.Tables.Count-1;

            Microsoft.Office.Interop.Excel.Application appExcel;
            appExcel = new Microsoft.Office.Interop.Excel.Application();
            
            Microsoft.Office.Interop.Excel.Workbook workbookData;
            Microsoft.Office.Interop.Excel.Worksheet worksheetData;

            workbookData = appExcel.Workbooks.Add(Missing.Value);
            //
            //workbookData.Worksheets.Delete();
            for(int k=0;k<ds.Tables.Count;k++)
            {
                worksheetData = (Microsoft.Office.Interop.Excel.Worksheet)workbookData.Worksheets.Add(Missing.Value,Missing.Value,Missing.Value,Missing.Value);
                worksheetData.Name = sheetName+"_"+testnum.ToString();
                testnum--;
                if(ds.Tables[k]!=null)
                {
                    for(int i=0;i<ds.Tables[k].Rows.Count;i++)
                    {
                        for(int j=0;j<ds.Tables[k].Columns.Count;j++)
                        {
                            worksheetData.Cells[i+1,j+1] = ds.Tables[k].Rows[i][j].ToString();
                        }
                    }
                }
                
                worksheetData.Columns.EntireColumn.AutoFit();
                workbookData.Saved = true;

            }
            //string strFileName = "C://Inetpub//wwwroot//External//Mongoose//files//"+ sheetName + ".xls";
            string strFileName = "e://www//页面//External//Mongoose//files//"+ sheetName + ".xls";
            
            workbookData.SaveCopyAs(strFileName);

            appExcel.Quit();

            //Response.Redirect("../Mongoose/files/"+sheetName+".xls");

        }
            

    }
}
