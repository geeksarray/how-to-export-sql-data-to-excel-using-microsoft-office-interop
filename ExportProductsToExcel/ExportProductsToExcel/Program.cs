using System;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Data.SqlClient;   
using System.Runtime.InteropServices;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;   

namespace ExportProductsToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string conString = "Data Source=AADITYA-PC;Initial Catalog=Northwind;Integrated Security=True";
            StringBuilder query = new StringBuilder();
            query.Append("SELECT Categories.CategoryName ");
            query.Append(",[ProductID], [ProductName], [SupplierID] ");
            query.Append(",[QuantityPerUnit], [UnitPrice], [UnitsInStock] ");
            query.Append(",[UnitsOnOrder], [ReorderLevel], [Discontinued] ");
            query.Append("FROM [northwind].[dbo].[Products] ");
            query.Append("JOIN Categories ON Categories.CategoryID = Products.CategoryID ");
            query.Append("ORDER BY Categories.CategoryName ");

            SQL.DataTable dtProducts = new SQL.DataTable(); 
            using (SqlConnection cn = new SqlConnection(conString))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query.ToString(), cn))
                {
                    da.Fill(dtProducts); 
                }
            }

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            
            oXL = new Excel.Application();
            oXL.Visible = true;

            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            
            try
            {
                SQL.DataTable dtCategories = dtProducts.DefaultView.ToTable(true, "CategoryName");

                foreach (SQL.DataRow category in dtCategories.Rows)
                {
                    oSheet = (Excel._Worksheet)oXL.Worksheets.Add();
                    oSheet.Name = category[0].ToString().Replace(" ", "").Replace("  ", "").Replace("/", "").Replace("\\", "").Replace("*", "");

                    string[] colNames = new string[dtProducts.Columns.Count];

                    int col = 0;

                    foreach (SQL.DataColumn dc in dtProducts.Columns)
                        colNames[col++] = dc.ColumnName;

                    char lastColumn = (char)(65 + dtProducts.Columns.Count - 1);

                    oSheet.get_Range("A1", lastColumn + "1").Value2 = colNames;
                    oSheet.get_Range("A1", lastColumn + "1").Font.Bold = true;
                    oSheet.get_Range("A1", lastColumn + "1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    SQL.DataRow[] dr = dtProducts.Select(string.Format("CategoryName='{0}'", category[0].ToString()));

                    string[,] rowData = new string[dr.Count<SQL.DataRow>(), dtProducts.Columns.Count];

                    int rowCnt = 0;
                    int redRows = 2;
                    foreach (SQL.DataRow row in dr)
                    {                         
                        for (col = 0; col < dtProducts.Columns.Count; col++)
                        {
                            rowData[rowCnt, col] = row[col].ToString();
                        }

                        if (int.Parse(row["ReorderLevel"].ToString()) < int.Parse(row["UnitsOnOrder"].ToString()))
                        {
                            Range range = oSheet.get_Range("A" + redRows.ToString(), "J" + redRows.ToString());
                            range.Cells.Interior.Color = System.Drawing.Color.Red; 
                        }
                        redRows++;
                        rowCnt++;
                    }
                    oSheet.get_Range("A2", lastColumn + rowCnt.ToString()).Value2 = rowData;
                }   

                oXL.Visible = true;
                oXL.UserControl = true;

                oWB.SaveAs("Products.xlsx",
                    AccessMode: Excel.XlSaveAsAccessMode.xlShared);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);  
            }
            finally
            {   
                Marshal.ReleaseComObject(oWB);
            }
        }
    }
}


