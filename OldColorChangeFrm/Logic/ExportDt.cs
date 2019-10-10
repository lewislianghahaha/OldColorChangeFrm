using System;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OldColorChangeFrm.DB;

namespace OldColorChangeFrm.Logic
{
    //导出
    public class ExportDt
    {
     //   DtList dtList=new DtList();

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAddress">导出地址</param>
        /// <param name="tempdt">运算结果-表头</param>
        public bool ExportDtToExcel(string fileAddress, DataTable tempdt/*, DataTable tempdtldt*/)
        {
            var result = true;
            var sheetcount = 0;  //记录所需的sheet页总数
            var rownum = 1;

            try
            {
                //声明一个WorkBook
                var xssfWorkbook = new XSSFWorkbook();
                //通过运算得出的表头及表体合并最终DT
                //Margedt(tempdt, tempdtldt);
                //执行sheet页(注:1)先列表tempdt行数判断需拆分多少个sheet表进行填充; 以一个sheet表有10W行记录填充为基准)
                sheetcount = tempdt.Rows.Count % 100000 == 0 ? tempdt.Rows.Count / 100000 : tempdt.Rows.Count / 100000 + 1;
                //i为EXCEL的Sheet页数ID
                for (var i = 1; i <= sheetcount; i++)
                {
                    //创建sheet页
                    var sheet = xssfWorkbook.CreateSheet("Sheet" + i);
                    //创建"标题行"
                    var row = sheet.CreateRow(0);
                    //创建sheet页各列标题
                    for (var j = 0; j < tempdt.Columns.Count; j++)
                    {
                        //设置列宽度
                        sheet.SetColumnWidth(j, (int)((20 + 0.72) * 256));
                        //创建标题
                        //0:横向导出方式
                        if (GlobalClasscs.ChooseType.ChooseTypeId == 0)
                        {
                            switch (j)
                            {
                                #region
                                case 0:
                                    row.CreateCell(j).SetCellValue("制造商");
                                    break;
                                case 1:
                                    row.CreateCell(j).SetCellValue("车型");
                                    break;
                                case 2:
                                    row.CreateCell(j).SetCellValue("涂层");
                                    break;
                                case 3:
                                    row.CreateCell(j).SetCellValue("颜色描述");
                                    break;
                                case 4:
                                    row.CreateCell(j).SetCellValue("内部色号");
                                    break;
                                case 5:
                                    row.CreateCell(j).SetCellValue("主配方色号(差异色)");
                                    break;
                                case 6:
                                    row.CreateCell(j).SetCellValue("颜色组别");
                                    break;
                                case 7:
                                    row.CreateCell(j).SetCellValue("标准色号");
                                    break;
                                case 8:
                                    row.CreateCell(j).SetCellValue("RGBValue");
                                    break;
                                case 9:
                                    row.CreateCell(j).SetCellValue("版本日期");
                                    break;
                                case 10:
                                    row.CreateCell(j).SetCellValue("层");
                                    break;
                                case 11:
                                    row.CreateCell(j).SetCellValue("制作人");
                                    break;
                                case 12:
                                    row.CreateCell(j).SetCellValue("旧系统配方号");
                                    break;
                                case 13:
                                    row.CreateCell(j).SetCellValue("色板来源");
                                    break;
                                case 14:
                                    row.CreateCell(j).SetCellValue("旧系统涂层");
                                    break;

                                case 15:
                                    row.CreateCell(j).SetCellValue("色母1");
                                    break;
                                case 16:
                                    row.CreateCell(j).SetCellValue("色母量1");
                                    break;
                                case 17:
                                    row.CreateCell(j).SetCellValue("色母2");
                                    break;
                                case 18:
                                    row.CreateCell(j).SetCellValue("色母量2");
                                    break;
                                case 19:
                                    row.CreateCell(j).SetCellValue("色母3");
                                    break;
                                case 20:
                                    row.CreateCell(j).SetCellValue("色母量3");
                                    break;
                                case 21:
                                    row.CreateCell(j).SetCellValue("色母4");
                                    break;
                                case 22:
                                    row.CreateCell(j).SetCellValue("色母量4");
                                    break;
                                case 23:
                                    row.CreateCell(j).SetCellValue("色母5");
                                    break;
                                case 24:
                                    row.CreateCell(j).SetCellValue("色母量5");
                                    break;
                                case 25:
                                    row.CreateCell(j).SetCellValue("色母6");
                                    break;
                                case 26:
                                    row.CreateCell(j).SetCellValue("色母量6");
                                    break;
                                case 27:
                                    row.CreateCell(j).SetCellValue("色母7");
                                    break;
                                case 28:
                                    row.CreateCell(j).SetCellValue("色母量7");
                                    break;
                                case 29:
                                    row.CreateCell(j).SetCellValue("色母8");
                                    break;
                                case 30:
                                    row.CreateCell(j).SetCellValue("色母量8");
                                    break;
                                case 31:
                                    row.CreateCell(j).SetCellValue("色母9");
                                    break;
                                case 32:
                                    row.CreateCell(j).SetCellValue("色母量9");
                                    break;
                                case 33:
                                    row.CreateCell(j).SetCellValue("色母10");
                                    break;
                                case 34:
                                    row.CreateCell(j).SetCellValue("色母量10");
                                    break;
                                case 35:
                                    row.CreateCell(j).SetCellValue("色母11");
                                    break;
                                case 36:
                                    row.CreateCell(j).SetCellValue("色母量11");
                                    break;

                                case 37:
                                    row.CreateCell(j).SetCellValue("色母12");
                                    break;
                                case 38:
                                    row.CreateCell(j).SetCellValue("色母量12");
                                    break;
                                case 39:
                                    row.CreateCell(j).SetCellValue("色母13");
                                    break;
                                case 40:
                                    row.CreateCell(j).SetCellValue("色母量13");
                                    break;
                                case 41:
                                    row.CreateCell(j).SetCellValue("色母14");
                                    break;
                                case 42:
                                    row.CreateCell(j).SetCellValue("色母量14");
                                    break;
                                case 43:
                                    row.CreateCell(j).SetCellValue("色母15");
                                    break;
                                case 44:
                                    row.CreateCell(j).SetCellValue("色母量15");
                                    break;
                                case 45:
                                    row.CreateCell(j).SetCellValue("色母16");
                                    break;
                                case 46:
                                    row.CreateCell(j).SetCellValue("色母量16");
                                    break;
                                case 47:
                                    row.CreateCell(j).SetCellValue("色母17");
                                    break;
                                case 48:
                                    row.CreateCell(j).SetCellValue("色母量17");
                                    break;
                                case 49:
                                    row.CreateCell(j).SetCellValue("色母18");
                                    break;
                                case 50:
                                    row.CreateCell(j).SetCellValue("色母量18");
                                    break;
                                case 51:
                                    row.CreateCell(j).SetCellValue("色母19");
                                    break;
                                case 52:
                                    row.CreateCell(j).SetCellValue("色母量19");
                                    break;
                                case 53:
                                    row.CreateCell(j).SetCellValue("色母20");
                                    break;
                                case 54:
                                    row.CreateCell(j).SetCellValue("色母量21");
                                    break;
                                    #endregion
                            }
                        }
                        //1:竖向导出方式
                        else
                        {
                            switch (j)
                            {
                                #region SetCellValue
                                case 0:
                                    row.CreateCell(j).SetCellValue("制造商");
                                    break;
                                case 1:
                                    row.CreateCell(j).SetCellValue("车型");
                                    break;
                                case 2:
                                    row.CreateCell(j).SetCellValue("涂层");
                                    break;
                                case 3:
                                    row.CreateCell(j).SetCellValue("颜色描述");
                                    break;
                                case 4:
                                    row.CreateCell(j).SetCellValue("内部色号");
                                    break;
                                case 5:
                                    row.CreateCell(j).SetCellValue("主配方色号(差异色)");
                                    break;
                                case 6:
                                    row.CreateCell(j).SetCellValue("颜色组别");
                                    break;
                                case 7:
                                    row.CreateCell(j).SetCellValue("标准色号");
                                    break;
                                case 8:
                                    row.CreateCell(j).SetCellValue("RGBValue");
                                    break;
                                case 9:
                                    row.CreateCell(j).SetCellValue("版本日期");
                                    break;
                                case 10:
                                    row.CreateCell(j).SetCellValue("层");
                                    break;
                                case 11:
                                    row.CreateCell(j).SetCellValue("色母编码");
                                    break;
                                case 12:
                                    row.CreateCell(j).SetCellValue("色母名称");
                                    break;
                                case 13:
                                    row.CreateCell(j).SetCellValue("色母量(克)");
                                    break;
                                case 14:
                                    row.CreateCell(j).SetCellValue("累积量");
                                    break;
                                case 15:
                                    row.CreateCell(j).SetCellValue("制作人");
                                    break;
                                case 16:
                                    row.CreateCell(j).SetCellValue("旧系统配方号");
                                    break;
                                case 17:
                                    row.CreateCell(j).SetCellValue("色板来源");
                                    break;
                                case 18:
                                    row.CreateCell(j).SetCellValue("旧系统涂层");
                                    break;
                                    #endregion
                            }
                        }
                    }

                    //计算进行循环的起始行
                    var startrow = (i - 1) * 100000;
                    //计算进行循环的结束行
                    var endrow = i == sheetcount ? tempdt.Rows.Count : i * 100000;

                    //每一个sheet表显示100000行  
                    for (var j = startrow; j < endrow; j++)
                    {
                        //创建行
                        row = sheet.CreateRow(rownum);
                        //循环获取DT内的列值记录
                        for (var k = 0; k < tempdt.Columns.Count; k++)
                        {
                            if(Convert.ToString(tempdt.Rows[j][k]) == "") continue;
                            else
                            {
                                //0:横向导出方式
                                if (GlobalClasscs.ChooseType.ChooseTypeId == 0)
                                {
                                    row.CreateCell(k, CellType.String).SetCellValue(Convert.ToString(tempdt.Rows[j][k]));
                                }
                                //1:竖向导出方式
                                else
                                {
                                    //当ColNum=13 或 14 时,执行(注:要注意值小数位数保留两位;当超出三位小数的时候,会出现OutofMemory异常.)
                                    if (k == 13 || k == 14)
                                    {
                                        row.CreateCell(k, CellType.Numeric).SetCellValue(Convert.ToDouble(tempdt.Rows[j][k]));
                                    }
                                    //除‘色母量’以及‘累积量’外的值的转换赋值
                                    else
                                    {
                                        row.CreateCell(k, CellType.String).SetCellValue(Convert.ToString(tempdt.Rows[j][k]));
                                    }
                                }
                            }
                        }
                        rownum++;
                    }
                    //当一个SHEET页填充完毕后,需将变量初始化
                    rownum = 1;
                }

                //写入数据
                var file = new FileStream(fileAddress, FileMode.Create);
                xssfWorkbook.Write(file);
                file.Close();           //关闭文件流
                xssfWorkbook.Close();   //关闭工作簿
                file.Dispose();         //释放文件流
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }
    }
}
