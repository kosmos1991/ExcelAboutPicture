using ExcelAboutPicture.ViewModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace ExcelAboutPicture.Models
{
    public class ImportExcelHelper
    {
        /// <summary>
        /// 生成EXCEL
        /// </summary>
        /// <param name="info">数据</param>
        /// <param name="sSheetName">EXCEL生成后的路径，绝对路径如：C:\a.xls</param>
        public ImportExcelHelper(List<DataListViewModel> info, string sSheetName)
        {
            FileStream fs = null;
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("测试列表"); //创建一个名称为XX的表
            sheet.CreateFreezePane(0, 1); //冻结列头行
            IRow row_Title = sheet.CreateRow(0); //创建列头行
            row_Title.HeightInPoints = 19.5F; //设置列头行高

            // 总列
            var row = 2;

            #region 设置列宽
            // 设置列宽,excel列宽每个像素是1/256
            for (int i = 0; i <= row; i++)
            {
                switch (i)
                {
                    case 0:
                        sheet.SetColumnWidth(i, 10 * 256);
                        break;
                    case 1:
                        sheet.SetColumnWidth(i, 20 * 256);
                        break;
                }
            }
            #endregion

            #region 设置列头单元格样式
            ICellStyle cs_Title = wb.CreateCellStyle(); //创建列头样式
            cs_Title.Alignment = HorizontalAlignment.Center; //水平居中
            cs_Title.VerticalAlignment = VerticalAlignment.Center; //垂直居中
            IFont cs_Title_Font = wb.CreateFont(); //创建字体
            cs_Title_Font.IsBold = true; //字体加粗
            cs_Title_Font.FontHeightInPoints = 12; //字体大小
            cs_Title.SetFont(cs_Title_Font); //将字体绑定到样式
            #endregion

            #region 生成列头
            for (int i = 0; i <= row; i++)
            {
                ICell cell_Title = row_Title.CreateCell(i); //创建单元格
                cell_Title.CellStyle = cs_Title; //将样式绑定到单元格
                switch (i)
                {
                    case 0:
                        cell_Title.SetCellValue("序号");
                        break;
                    case 1:
                        cell_Title.SetCellValue("图片");
                        break;
                }
            }
            #endregion

            if (info != null)
            {
                for (int i = 0; i < info.Count; i++)
                {
                    #region 设置内容单元格样式
                    ICellStyle cs_Content = wb.CreateCellStyle(); //创建列头样式
                    cs_Content.Alignment = HorizontalAlignment.Center; //水平居中
                    cs_Content.VerticalAlignment = VerticalAlignment.Center; //垂直居中
                    cs_Content.WrapText = true; // 自动换行
                    IFont cs_Content_Font = wb.CreateFont(); //创建字体
                    cs_Content_Font.FontHeightInPoints = 12; //字体大小
                    cs_Content.SetFont(cs_Content_Font); //将字体绑定到样式
                    #endregion

                    #region 生成内容单元格
                    //创建行
                    IRow row_Content = sheet.CreateRow(i + 1);
                    //设置行高 ,excel行高度每个像素点是1/20
                    row_Content.Height = 100 * 20;
                    for (int j = 0; j <= row; j++)
                    {
                        ICell cell_Conent = row_Content.CreateCell(j); //创建单元格
                        cell_Conent.CellStyle = cs_Content;
                        switch (j)
                        {
                            case 0:
                                SetCellValue(cell_Conent, info[i].ID);
                                break;
                            case 1:
                                if (!string.IsNullOrEmpty(info[i].ImgPath))
                                {
                                    // 第一步：读取图片到byte数组
                                    var filename = GetImageRoute(info[i].ImgPath);
                                    if (!string.IsNullOrEmpty(filename))
                                    {
                                        byte[] bytes = File.ReadAllBytes(filename);

                                        // 第二步：将图片添加到workbook中 指定图片格式 返回图片所在workbook->Picture数组中的索引地址（从1开始）
                                        int pictureIdx = wb.AddPicture(bytes, PictureType.JPEG);

                                        // 第三步：在sheet中创建画部
                                        IDrawing drawing = sheet.CreateDrawingPatriarch();

                                        // 第四步：设置锚点 （在起始单元格的X坐标0-1023，Y的坐标0-255，在终止单元格的X坐标0-1023，Y的坐标0-255，起始单元格行数，列数，终止单元格行数，列数）
                                        IClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, j, i + 1, j + 1, i + 2);

                                        // 第五步：创建图片
                                        IPicture picture = drawing.CreatePicture(anchor, pictureIdx);

                                        // 删除本地图片
                                        DeleteImage(filename);
                                    }
                                }
                                break;
                        }
                    }
                    #endregion
                }
            }

            using (fs = File.OpenWrite(sSheetName))
            {
                wb.Write(fs);//向打开的这个xls文件中写入数据
            }
        }

        /// <summary>
        /// 根据数据类型设置不同类型的cell
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="obj"></param>
        public static void SetCellValue(ICell cell, object obj)
        {
            try
            {
                if (obj.GetType() == typeof(int))
                {
                    cell.SetCellValue((int)obj);
                }
                else if (obj.GetType() == typeof(double))
                {
                    cell.SetCellValue((double)obj);
                }
                else if (obj.GetType() == typeof(string))
                {
                    cell.SetCellValue((string)obj);
                }
                else if (obj.GetType() == typeof(string))
                {
                    cell.SetCellValue(obj.ToString());
                }
                else if (obj.GetType() == typeof(DateTime))
                {
                    cell.SetCellValue(Convert.ToDateTime(obj).ToString("yyyy/MM/dd hh:mm:ss"));
                }
                else if (obj.GetType() == typeof(bool))
                {
                    cell.SetCellValue((bool)obj);
                }
                else
                {
                    cell.SetCellValue(obj.ToString());
                }
            }
            catch (Exception ex)
            {
            }
        }

        #region 从图片地址下载图片到本地磁盘
        /// <summary>
        /// 获取图片的本地路径
        /// </summary>
        /// <param name="picUrl">远程url路径</param>
        /// <returns></returns>
        public static string GetImageRoute(string picUrl)
        {
            try
            {
                var Folder = AppDomain.CurrentDomain.BaseDirectory + "/Upload";

                if (Directory.Exists(Folder) == false)//如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(Folder);
                }

                var fileName = Folder + "/" + Guid.NewGuid().ToString().Replace("-", "") + ".jpg";
                DownloadPicture(picUrl, fileName);
                return fileName;
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        /// <summary>
        /// 从图片地址下载图片到本地磁盘
        /// </summary>
        /// <param name="picUrl">远程url路径</param>
        /// <param name="savePath">本地路径</param>
        /// <param name="timeOut"></param>
        /// <returns></returns>
        private static bool DownloadPicture(string picUrl, string savePath, int timeOut = -1)
        {
            bool value = false;
            WebResponse response = null;
            Stream stream = null;
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(picUrl);
                if (timeOut != -1) request.Timeout = timeOut;
                response = request.GetResponse();
                stream = response.GetResponseStream();
                if (!response.ContentType.ToLower().StartsWith("text/"))
                    value = SaveBinaryFile(response, savePath);
            }
            finally
            {
                if (stream != null) stream.Close();
                if (response != null) response.Close();
            }
            return value;
        }

        private static bool SaveBinaryFile(WebResponse response, string savePath)
        {
            bool value = false;
            byte[] buffer = new byte[1024];
            Stream outStream = null;
            Stream inStream = null;
            try
            {
                if (File.Exists(savePath)) File.Delete(savePath);
                outStream = System.IO.File.Create(savePath);
                inStream = response.GetResponseStream();
                int l;
                do
                {
                    l = inStream.Read(buffer, 0, buffer.Length);
                    if (l > 0) outStream.Write(buffer, 0, l);
                } while (l > 0);
                value = true;
            }
            finally
            {
                if (outStream != null) outStream.Close();
                if (inStream != null) inStream.Close();
            }
            return value;
        }
        #endregion

        #region 删除本地图片
        public static void DeleteImage(string path)
        {
            File.Delete(path);
        }
        #endregion
    }
}