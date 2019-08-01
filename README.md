# ExcelAboutPicture
导出Excel之单元格插入图片

# 插件
安装NPOI官方插件：Install-Package NPOI

# 导出Excel步骤
1. 引入NPOI
2. 定义文件名和路径，路径如不存在则创建
3. 获取要导出的数据列表
4. 绘制Excel
5. 在指定路径生成Excel并返回Excel路径
6. 下载Excel

# 导出Excel之单元格插入图片步骤
1. 下载远程图片到本地
2. 将本地图片转为byte格式
3. 将图片添加到workbook中 指定图片格式 返回图片所在workbook->Picture数组中的索引地址（从1开始）
4. 在sheet中创建画部
5. 设置锚点 （在起始单元格的X坐标0-1023，Y的坐标0-255，在终止单元格的X坐标0-1023，Y的坐标0-255，起始单元格行数，列数，终止单元格行数，列数）
6. 创建图片
7. 删除本地图片

# 代码解析
https://blog.csdn.net/qq_31267183/article/details/97785772
