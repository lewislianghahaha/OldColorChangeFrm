using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OldColorChangeFrm.DB
{
    public class DtList
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ExportDt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 18; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String"); 
                        break;
                    case 1:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "主配方色号(差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "色母编码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 12:
                        dc.ColumnName = "色母名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 13:
                        dc.ColumnName = "色母量(克)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 14:
                        dc.ColumnName = "累积量";
                        dc.DataType = Type.GetType("System.Double"); 
                        break;
                    case 15:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 16:
                        dc.ColumnName = "旧系统配方号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 17:
                        dc.ColumnName = "色板来源";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }
    }
}
