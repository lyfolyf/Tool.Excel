using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lead.Tool.Excel
{
    public class ExcelHelper
    {
        static string LisencePath = @"D:\Plat.zzc\Bin\3rdLibs\License.lic";

        public static  List<string> Read(string path)
        {
            List<string> Re = new List<string>();
            try
            {
                //注册license
                {
                    //读取 License 文件
                    Stream stream = (Stream)File.OpenRead(LisencePath);

                    //注册 License
                    Aspose.Cells.License li = new Aspose.Cells.License();
                    li.SetLicense(stream);
                }

            
                Workbook workbook = new Workbook();
                workbook.Open(path);
                Cells cells = workbook.Worksheets[0].Cells;

                foreach (var item in cells)
                {
                    string Single = "";
                    for (int i = 0; i < cells.MaxDataRow + 1; i++)
                    {
                        for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                        {
                            string s = cells[i, j].StringValue.Trim();
                            Single += s + ",";
                        }
                    }
                    
                    Re.Add(Single);
                }
                
            }
            catch
            {
               ;
            }
            return Re;
        }

        public static void WriteOneLine(string FilePath, List< string> headStr, List<string> RowsData)
        {
            try
            {
                //注册license
                {
                    //读取 License 文件
                    Stream stream = (Stream)File.OpenRead(LisencePath);

                    //注册 License
                    Aspose.Cells.License li = new Aspose.Cells.License();
                    li.SetLicense(stream);
                }

                //新建文件
                {
                    if (!File.Exists(FilePath))
                    {
                        #region 创建文件
                        if (!File.Exists(FilePath))
                        {
                            File.Create(FilePath); //创建文件路径               
                        }

                        //设置文件访问权限，ReadWrite可读写，FileShare.Read 允许其他同时读取
                        //创建一个工作簿
                        Aspose.Cells.Workbook workbook11 = new Aspose.Cells.Workbook();

                        //创建一个 sheet 表
                        Aspose.Cells.Worksheet worksheet11 = workbook11.Worksheets[0];

                        //设置 sheet 表名称
                        worksheet11.Name = "sheet1";

                        Aspose.Cells.Cell cell0;

                        //设置列名
                        for (int i = 0; i < headStr.Count; i++)
                        {
                            {
                                //获取第一行的每个单元格
                                cell0 = worksheet11.Cells[0, 0 + i];

                                //设置列名
                                cell0.PutValue(headStr[i]);

                                //设置字体
                                cell0.Style.Font.Name = "Arial";

                                //设置字体加粗
                                cell0.Style.Font.IsBold = true;

                                //设置字体大小
                                cell0.Style.Font.Size = 12;

                                //设置字体颜色
                                cell0.Style.Font.Color = System.Drawing.Color.Black;

                                //设置背景色
                                cell0.Style.Pattern = BackgroundType.Solid;
                                cell0.Style.ForegroundColor = Color.YellowGreen;

                            } 
                        }

                        //保存至指定路径
                        workbook11.Save(FilePath);
                        worksheet11 = null;
                        workbook11 = null;
                        #endregion
                    }
                }

                {
                    //打开一个 excel 
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
                    workbook.Open(FilePath);

                    //打开一个 sheet 表
                    Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];


                    var in1 = worksheet.Index;

                    var temp = worksheet.Cells.MaxDataRow;

                    for (int i = 0; i < RowsData.Count; i++)
                    {
                        worksheet.Cells[temp + 1, i].PutValue(RowsData[i]);
                        worksheet.Cells[temp + 1, i].Style.Pattern = BackgroundType.Solid;
                        worksheet.Cells[temp + 1, i].Style.ForegroundColor = Color.White;

                    }

                    //自动列宽
                    // worksheet.AutoFitColumns();
                    workbook.Save(FilePath);
                    //string filename1 = string.Format(@"{0}{1}{2}.xls", FilePath, "Result", DateTime.Now.ToString("yyyyMMdd-1"));
                    //File.Copy(filename, filename1, true);
                    worksheet = null;
                    workbook = null;
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
