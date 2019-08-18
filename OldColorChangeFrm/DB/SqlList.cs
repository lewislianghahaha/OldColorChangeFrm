﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OldColorChangeFrm.DB
{
    public class SqlList
    {
        //根据SQLID返回对应的SQL语句  
        private string _result;

        public string Get_Result()
        {
            _result = $@"
                           SELECT a.fformulacode 配方代码,d.fname 品牌,b.fcolorid 颜色代码,b.fcolorName 颜色名称,c.fchname 车厂,
	                               b.fmatchModel 适用车型,a.fvariant 差异色,b.fcoat 涂层,a.fyear 年份,
	                               a.fsource 色板来源,a.fmfdate 制作日期,a.fmfname 制作人,
                                   f.fnbr 色母代码,f.fname1 色母名称,e.fqty '色母量(克)' 
                            FROM ColorScheme a
                            INNER JOIN dbo.ModeColor b ON a.fmodelColorId=b.fid
                            INNER JOIN dbo.CarFactory c ON b.fcarFactoryId=c.fid
                            INNER JOIN dbo.Brand d ON a.fbrandID=d.fid

                            INNER JOIN dbo.ColorSchemeDetail e ON a.fid=e.fcolorSchemeId
                            INNER JOIN dbo.Color f ON e.fcolorId=f.fid

                            WHERE d.fname='施莱威'--a.fformulacode='034215'
                            ORDER BY b.fcolorid,a.fformulacode 
                        ";
            return _result;
        }

    }
}
