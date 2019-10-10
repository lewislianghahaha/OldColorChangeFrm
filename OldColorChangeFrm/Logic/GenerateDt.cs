using System;
using System.Data;
using System.Data.SqlClient;
using CroMaxChangeFrm;
using OldColorChangeFrm.DB;

namespace OldColorChangeFrm.Logic
{
    //运算
    public class GenerateDt
    {
        DtList dtList=new DtList();
        SqlList sqlList=new SqlList();

        /// <summary>
        /// 运算-通过从EXCEL导入的DT获取表头信息
        /// </summary>
        /// <returns></returns>
        public DataTable Generatetemp()
        {
            var resultdt=new DataTable();

            //保存‘配方代码’字段,用于排除重复值
            var colorcode = string.Empty;

            try
            {
                //从数据库内获取的DT
                var sourcedt = GetSourceDt();
                //获取表头临时表(0:横向 1:竖向)
                resultdt = GlobalClasscs.ChooseType.ChooseTypeId == 0 ? dtList.Get_ExportVDt() : dtList.Get_ExportDt();
                //先循环从SQL内获取的DT
                foreach (DataRow rows in sourcedt.Rows)
                {
                    //若循环获取的‘配方代码’与变量一致,即不用继续
                    if (colorcode == Convert.ToString(rows[0])) continue;
                    //若不相同,先将当前循环行的值进行赋值至变量
                    colorcode = Convert.ToString(rows[0]);

                    //0:横向导出方式 1:竖向导出方式
                    resultdt.Merge(GlobalClasscs.ChooseType.ChooseTypeId == 0
                        ? GetVdt(rows,sourcedt,resultdt) : GetHdt(rows, sourcedt, resultdt));
                }
            }
            catch (Exception)
            {
                resultdt.Rows.Clear();
                resultdt.Columns.Clear();
            }

            return resultdt;
        }

        /// <summary>
        /// 竖向导出方式使用
        /// </summary>
        /// <returns></returns>
        private DataTable GetHdt(DataRow rows,DataTable sourcedt,DataTable resultdt)
        {
            var rowsdtl = sourcedt.Select("配方代码='" + Convert.ToString(rows[0]) + "'");

            for (var i = 0; i < rowsdtl.Length; i++)
            {
                var newrows = resultdt.NewRow();
                newrows[0] = i == 0 ? rowsdtl[i][4] : DBNull.Value;      //制造商
                newrows[1] = i == 0 ? rowsdtl[i][5] : DBNull.Value;      //车型
                newrows[2] = i == 0 ? rowsdtl[i][7] : DBNull.Value;      //涂层
                newrows[3] = i == 0 ? rowsdtl[i][3] : DBNull.Value;      //颜色描述
                newrows[4] = "";                                         //内部色号
                newrows[5] = i == 0 ? rowsdtl[i][6] : DBNull.Value;      //主配方色号(差异色)
                newrows[6] = "";                                         //颜色组别
                newrows[7] = i == 0 ? rowsdtl[i][2] : DBNull.Value;      //标准色号
                newrows[8] = i == 0 ? "#9e5014" : "";                    //RGBValue
                newrows[9] = i == 0 ? rowsdtl[i][10] : DBNull.Value;     //版本日期
                newrows[10] = "";                                        //层

                newrows[11] = rowsdtl[i][12];                                   //色母编码
                newrows[12] = rowsdtl[i][13];                                   //色母名称
                newrows[13] = Math.Round(Convert.ToDouble(rowsdtl[i][14]), 2);  //色母量(克)
                newrows[14] = DBNull.Value;                                     //累积量

                newrows[15] = i == 0 ? rowsdtl[i][11] : DBNull.Value;    //制作人
                newrows[16] = i == 0 ? rowsdtl[i][0] : DBNull.Value;     //旧系统配方号
                newrows[17] = i == 0 ? rowsdtl[i][9] : DBNull.Value;     //色板来源
                newrows[18] = i == 0 ? rowsdtl[i][7] : DBNull.Value;     //旧系统涂层

                //对‘涂层’及‘层’根据指定条件进行修改赋值
                switch (Convert.ToString(newrows[2]))
                {
                    case "":
                        newrows[2] = "";
                        newrows[10] = "";
                        break;
                    case "底色漆":
                        newrows[2] = "两工序";
                        newrows[10] = "1";
                        break;
                    case "面色漆":
                        newrows[2] = "素色";
                        newrows[10] = "1";
                        break;
                    case "3工序-底漆":
                        newrows[2] = "三工序";
                        newrows[10] = "1";
                        break;
                    case "3工序-面漆":
                        newrows[2] = "三工序";
                        newrows[10] = "2";
                        break;
                }
                resultdt.Rows.Add(newrows);
            }
            return resultdt;
        }

        /// <summary>
        /// 横向导出方式使用
        /// </summary>
        /// <returns></returns>
        private DataTable GetVdt(DataRow rows, DataTable sourcedt, DataTable resultdt)
        {
            //先将‘制造商’等表头相关信息插入,再插入色母等信息
            var newrow = resultdt.NewRow();
            newrow[0] = rows[4];             //制造商
            newrow[1] = rows[5];             //车型
            newrow[2] = "";                  //涂层
            newrow[3] = rows[3];             //颜色描述
            newrow[4] = "";                  //内部色号
            newrow[5] = rows[6];             //主配方色号(差异色)
            newrow[6] = "";                  //颜色组别
            newrow[7] = rows[2];             //标准色号
            newrow[8] = "#9e5014";           //RGBValue
            newrow[9] = rows[10];            //版本日期
            newrow[10] = "";                 //层
            newrow[11] = rows[11];           //制作人
            newrow[12] = rows[0];            //旧系统配方号
            newrow[13] = rows[9];            //色板来源
            newrow[14] = rows[7];            //旧系统涂层

            //对‘涂层’及‘层’根据指定条件进行修改赋值
            switch (Convert.ToString(newrow[2]))
            {
                case "":
                    newrow[2] = "";
                    newrow[10] = "";
                    break;
                case "底色漆":
                    newrow[2] = "两工序";
                    newrow[10] = "1";
                    break;
                case "面色漆":
                    newrow[2] = "素色";
                    newrow[10] = "1";
                    break;
                case "3工序-底漆":
                    newrow[2] = "三工序";
                    newrow[10] = "1";
                    break;
                case "3工序-面漆":
                    newrow[2] = "三工序";
                    newrow[10] = "2";
                    break;
            }

            //将‘色母’相关信息，插入至对应的项内
            var rowsdtl = sourcedt.Select("配方代码='" + Convert.ToString(rows[0]) + "'");

            for (var i = 0; i < rowsdtl.Length; i++)
            {
                newrow[15+i+i] = rowsdtl[i][12]; //色母编码
                newrow[15+i+i+1] = Math.Round(Convert.ToDouble(rowsdtl[i][14]), 2); //色母量(保留两位小数)
            }
            resultdt.Rows.Add(newrow);

            return resultdt;
        }

        /// <summary>
        /// 获取数据源
        /// </summary>
        /// <returns></returns>
        private DataTable GetSourceDt()
        {

            var dt = new DataTable();

            try
            {
                var sqlscript = sqlList.Get_Result();
                var sqlDataAdapter = new SqlDataAdapter(sqlscript, GetConn());
                sqlDataAdapter.Fill(dt);
            }
            catch (Exception)
            {
                dt.Columns.Clear();
                dt.Rows.Clear();
            }
            return dt;
        }

        /// <summary>
        /// 获取连接
        /// </summary>
        /// <returns></returns>
        public SqlConnection GetConn()
        {
            var conn = new Conn();
            var sqlcon = new SqlConnection(conn.GetConnectionString());
            return sqlcon;
        }
    }
}
