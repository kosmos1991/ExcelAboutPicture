using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelAboutPicture.ViewModel
{
    /// <summary>
    /// 数据视图
    /// </summary>
    public class DataListViewModel
    {
        /// <summary>
        /// 主键
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// 图片地址
        /// </summary>
        public string ImgPath { get; set; }
    }
}