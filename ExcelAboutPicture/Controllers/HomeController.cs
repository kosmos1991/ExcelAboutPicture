using ExcelAboutPicture.Models;
using ExcelAboutPicture.ViewModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelAboutPicture.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "首页";

            return View();
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <returns></returns>
        public JsonResult Export()
        {
            // 导出
            var now = DateTime.Now;
            var time = now.ToString("yyyyMMddHHmmss");
            var name = @"/Upload/";

            string namePath = AppDomain.CurrentDomain.BaseDirectory + name;//获取上传路径的物理地址 
            if (!Directory.Exists(namePath))//判断文件夹是否存在 
            {
                Directory.CreateDirectory(namePath);//不存在则创建文件夹 
            }

            name += "测试导出Excel" + time + ".xlsx";

            #region 生成数据
            List<DataListViewModel> list = new List<DataListViewModel>();
            list.Add(new DataListViewModel() { ID = 1, ImgPath = "http://pic41.nipic.com/20140508/18609517_112216473140_2.jpg" });
            list.Add(new DataListViewModel() { ID = 2, ImgPath = "http://pic31.nipic.com/20130801/11604791_100539834000_2.jpg" });
            #endregion

            ImportExcelHelper import = new ImportExcelHelper(list, AppDomain.CurrentDomain.BaseDirectory + name);

            var jsonResult = new
            {
                path = name,
            };

            return Json(jsonResult);
        }
    }
}
