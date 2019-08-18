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
                //获取表头临时表
                resultdt = dtList.Get_ExportDt();
                //先循环从SQL内获取的DT
                foreach (DataRow rows in sourcedt.Rows)
                {
                    //若循环获取的‘配方代码’与变量一致,即不用继续
                    if(colorcode==Convert.ToString(rows[0])) continue;
                    //若不相同,先将当前循环行的值进行赋值至变量
                    colorcode = Convert.ToString(rows[0]);
                    //再根据‘配方代码’为条件,放到数据源内进行查询及循环赋值
                    var rowsdtl = sourcedt.Select("配方代码='" + Convert.ToString(rows[0]) + "'");

                    for (var i = 0; i < rowsdtl.Length; i++)
                    {
                        var newrows = resultdt.NewRow();
                        newrows[0] = i == 0 ? rowsdtl[i][4] : DBNull.Value;      //制造商
                        newrows[1] = i == 0 ? rowsdtl[i][5] : DBNull.Value;      //车型
                        newrows[2] = i == 0 ? rowsdtl[i][7] : DBNull.Value;      //涂层
                        newrows[3] = i == 0 ? rowsdtl[i][3] : DBNull.Value;      //颜色描述
                        newrows[4] = "";                                         //内部色号
                        newrows[5] = "";                                         //主配方色号(差异色)
                        newrows[6] = "";                                         //颜色组别
                        newrows[7] = i == 0 ? rowsdtl[i][2] : DBNull.Value;      //标准色号
                        newrows[8] = i == 0 ? "#9e5014" : "";                    //RGBValue
                        newrows[9] = i == 0 ? rowsdtl[i][10] : DBNull.Value;     //版本日期
                        newrows[10] ="";                                         //层
                        newrows[11] =rowsdtl[i][12];                             //色母编码
                        newrows[12] =rowsdtl[i][13];                             //色母名称
                        newrows[13] =Math.Round(Convert.ToDouble(rowsdtl[i][14]),2);  //色母量(克)
                        newrows[14] =DBNull.Value;                               //累积量
                        newrows[15] = i == 0 ? rowsdtl[i][11] : DBNull.Value;    //制作人
                        newrows[16] = i == 0 ? rowsdtl[i][0] : DBNull.Value;     //旧系统配方号

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
