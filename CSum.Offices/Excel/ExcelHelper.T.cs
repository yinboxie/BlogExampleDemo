using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web;
using NPOI.HPSF;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace CSum.Offices.Excel
{
    /// <summary>
    ///     描 述：NPOI Excel泛型操作类
    public class ExcelHelper<T>
    {
        #region Excel导出方法 ExcelDownload

        /// <summary>
        ///     Excel导出下载
        /// </summary>
        /// <param name="lists">List<T>数据源</param>
        /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
        public static void ExcelDownload(List<T> lists, ExcelConfig excelConfig)
        {
            var curContext = HttpContext.Current;
            // 设置编码和附件格式
            curContext.Response.ContentType = "application/ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            curContext.Response.AppendHeader("Content-Disposition",
                "attachment;filename=" + HttpUtility.UrlEncode(excelConfig.FileName, Encoding.UTF8));
            //调用导出具体方法Export()
            curContext.Response.BinaryWrite(ExportMemoryStream(lists, excelConfig).GetBuffer());
            curContext.Response.End();
        }

        #endregion

        #region DataTable导出到Excel文件excelConfig中FileName设置为全路径

        /// <summary>
        ///     List<T>导出到Excel文件 ExcelImport
        /// </summary>
        /// <param name="lists">List<T>数据源</param>
        /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
        public static void ExcelImport(List<T> lists, ExcelConfig excelConfig)
        {
            using (var ms = ExportMemoryStream(lists, excelConfig))
            {
                using (var fs = new FileStream(excelConfig.FileName, FileMode.Create, FileAccess.Write))
                {
                    var data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }

        #endregion

        #region DataTable导出到Excel的MemoryStream

        /// <summary>
        ///     DataTable导出到Excel的MemoryStream Export()
        /// </summary>
        /// <param name="dtSource">DataTable数据源</param>
        /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
        public static MemoryStream ExportMemoryStream(List<T> lists, ExcelConfig excelConfig)
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet();
            var type = typeof(T);
            var properties = type.GetProperties();

            #region 右击文件 属性信息

            {
                var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "NPOI";
                workbook.DocumentSummaryInformation = dsi;

                var si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "刘晓雷"; //填加xls文件作者信息
                si.ApplicationName = "力软信息"; //填加xls文件创建程序信息
                si.LastAuthor = "刘晓雷"; //填加xls文件最后保存者信息
                si.Comments = "刘晓雷"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "主题信息"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion

            #region 设置标题样式

            var headStyle = workbook.CreateCellStyle();
            var arrColWidth = new int[properties.Length];
            var arrColName = new string[properties.Length]; //列名
            var arryColumStyle = new ICellStyle[properties.Length]; //样式表
            headStyle.Alignment = HorizontalAlignment.Center; // ------------------
            if (excelConfig.Background != new Color())
                if (excelConfig.Background != new Color())
                {
                    headStyle.FillPattern = FillPattern.SolidForeground;
                    headStyle.FillForegroundColor = GetXLColour(workbook, excelConfig.Background);
                }

            var font = workbook.CreateFont();
            font.FontHeightInPoints = excelConfig.TitlePoint;
            if (excelConfig.ForeColor != new Color()) font.Color = GetXLColour(workbook, excelConfig.ForeColor);
            font.Boldweight = 700;
            headStyle.SetFont(font);

            #endregion

            #region 列头及样式

            var cHeadStyle = workbook.CreateCellStyle();
            cHeadStyle.Alignment = HorizontalAlignment.Center; // ------------------
            var cfont = workbook.CreateFont();
            cfont.FontHeightInPoints = excelConfig.HeadPoint;
            cHeadStyle.SetFont(cfont);

            #endregion

            #region 设置内容单元格样式

            var i = 0;
            foreach (var column in properties)
            {
                var columnStyle = workbook.CreateCellStyle();
                columnStyle.Alignment = HorizontalAlignment.Center;
                arrColWidth[i] = Encoding.GetEncoding(936).GetBytes(column.Name).Length;
                arrColName[i] = column.Name;

                if (excelConfig.ColumnEntity != null)
                {
                    var columnentity = excelConfig.ColumnEntity.Find(t => t.Column == column.Name);
                    if (columnentity != null)
                    {
                        arrColName[i] = columnentity.ExcelColumn;
                        if (columnentity.Width != 0) arrColWidth[i] = columnentity.Width;
                        if (columnentity.Background != new Color())
                            if (columnentity.Background != new Color())
                            {
                                columnStyle.FillPattern = FillPattern.SolidForeground;
                                columnStyle.FillForegroundColor = GetXLColour(workbook, columnentity.Background);
                            }

                        if (columnentity.Font != null || columnentity.Point != 0 ||
                            columnentity.ForeColor != new Color())
                        {
                            var columnFont = workbook.CreateFont();
                            columnFont.FontHeightInPoints = 10;
                            if (columnentity.Font != null) columnFont.FontName = columnentity.Font;
                            if (columnentity.Point != 0) columnFont.FontHeightInPoints = columnentity.Point;
                            if (columnentity.ForeColor != new Color())
                                columnFont.Color = GetXLColour(workbook, columnentity.ForeColor);
                            columnStyle.SetFont(font);
                        }
                    }
                }

                arryColumStyle[i] = columnStyle;
                i++;
            }

            #endregion

            #region 填充数据

            #endregion

            var dateStyle = workbook.CreateCellStyle();
            var format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            var rowIndex = 0;
            foreach (var item in lists)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0) sheet = workbook.CreateSheet();

                    #region 表头及样式

                    {
                        if (excelConfig.Title != null)
                        {
                            var headerRow = sheet.CreateRow(0);
                            if (excelConfig.TitleHeight != 0) headerRow.Height = (short) (excelConfig.TitleHeight * 20);
                            headerRow.HeightInPoints = 25;
                            headerRow.CreateCell(0).SetCellValue(excelConfig.Title);
                            headerRow.GetCell(0).CellStyle = headStyle;
                            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, lists.Count - 1)); // ------------------
                        }
                    }

                    #endregion

                    #region 列头及样式

                    {
                        var headerRow = sheet.CreateRow(1);

                        #region 如果设置了列标题就按列标题定义列头，没定义直接按字段名输出

                        var headIndex = 0;
                        foreach (var column in properties)
                        {
                            headerRow.CreateCell(headIndex).SetCellValue(arrColName[headIndex]);
                            headerRow.GetCell(headIndex).CellStyle = cHeadStyle;
                            //设置列宽
                            sheet.SetColumnWidth(headIndex, (arrColWidth[headIndex] + 1) * 256);
                            headIndex++;
                        }

                        #endregion
                    }

                    #endregion

                    rowIndex = 2;
                }

                #endregion

                #region 填充内容

                var dataRow = sheet.CreateRow(rowIndex);
                var ordinal = 0;
                foreach (var column in properties)
                {
                    var newCell = dataRow.CreateCell(ordinal);
                    newCell.CellStyle = arryColumStyle[ordinal];
                    var drValue = column.GetValue(item, null) == null ? "" : column.GetValue(item, null).ToString();
                    SetCell(newCell, dateStyle, column.PropertyType, drValue);
                    ordinal++;
                }

                #endregion

                rowIndex++;
            }

            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }
        }

        #endregion

        #region 设置表格内容

        private static void SetCell(ICell newCell, ICellStyle dateStyle, Type dataType, string drValue)
        {
            switch (dataType.ToString())
            {
                case "System.String": //字符串类型
                    newCell.SetCellValue(drValue);
                    break;
                case "System.DateTime": //日期类型
                    DateTime dateV;
                    if (DateTime.TryParse(drValue, out dateV))
                        newCell.SetCellValue(dateV);
                    else
                        newCell.SetCellValue("");
                    newCell.CellStyle = dateStyle; //格式化显示
                    break;
                case "System.Boolean": //布尔型
                    var boolV = false;
                    bool.TryParse(drValue, out boolV);
                    newCell.SetCellValue(boolV);
                    break;
                case "System.Int16": //整型
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    var intV = 0;
                    int.TryParse(drValue, out intV);
                    newCell.SetCellValue(intV);
                    break;
                case "System.Decimal": //浮点型
                case "System.Double":
                    double doubV = 0;
                    double.TryParse(drValue, out doubV);
                    newCell.SetCellValue(doubV);
                    break;
                case "System.DBNull": //空值处理
                    newCell.SetCellValue("");
                    break;
                default:
                    newCell.SetCellValue("");
                    break;
            }
        }

        #endregion

        #region 读取excel ,默认第一行为标头

        /// <summary>
        ///     导入Excel
        /// </summary>
        /// <param name="lists"></param>
        /// <param name="head">中文列名对照</param>
        /// <param name="workbookFile">Excel所在路径</param>
        /// <returns></returns>
        public List<T> ExcelImport(Hashtable head, string workbookFile)
        {
            try
            {
                HSSFWorkbook hssfworkbook;
                var lists = new List<T>();
                using (var file = new FileStream(workbookFile, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new HSSFWorkbook(file);
                }

                var sheet = hssfworkbook.GetSheetAt(0) as HSSFSheet;
                var rows = sheet.GetRowEnumerator();
                var headerRow = sheet.GetRow(0) as HSSFRow;
                int cellCount = headerRow.LastCellNum;
                //Type type = typeof(T);
                PropertyInfo[] properties;
                var t = default(T);
                for (var i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i) as HSSFRow;
                    t = Activator.CreateInstance<T>();
                    properties = t.GetType().GetProperties();
                    foreach (var column in properties)
                    {
                        var j = headerRow.Cells.FindIndex(delegate(ICell c)
                        {
                            return c.StringCellValue == (head[column.Name] == null
                                       ? column.Name
                                       : head[column.Name].ToString());
                        });
                        if (j >= 0 && row.GetCell(j) != null)
                        {
                            var value = valueType(column.PropertyType, row.GetCell(j).ToString());
                            column.SetValue(t, value, null);
                        }
                    }

                    lists.Add(t);
                }

                return lists;
            }
            catch (Exception ee)
            {
                var see = ee.Message;
                return null;
            }
        }

        #endregion

        #region RGB颜色转NPOI颜色

        private static short GetXLColour(HSSFWorkbook workbook, Color SystemColour)
        {
            short s = 0;
            var XlPalette = workbook.GetCustomPalette();
            var XlColour = XlPalette.FindColor(SystemColour.R, SystemColour.G, SystemColour.B);
            if (XlColour == null)
            {
                if (PaletteRecord.STANDARD_PALETTE_SIZE < 255)
                {
                    XlColour = XlPalette.FindSimilarColor(SystemColour.R, SystemColour.G, SystemColour.B);
                    s = XlColour.Indexed;
                }
            }
            else
            {
                s = XlColour.Indexed;
            }

            return s;
        }

        #endregion

        private object valueType(Type t, string value)
        {
            object o = null;
            var strt = "String";
            if (t.Name == "Nullable`1") strt = t.GetGenericArguments()[0].Name;
            switch (strt)
            {
                case "Decimal":
                    o = decimal.Parse(value);
                    break;
                case "Int":
                    o = int.Parse(value);
                    break;
                case "Float":
                    o = float.Parse(value);
                    break;
                case "DateTime":
                    o = DateTime.Parse(value);
                    break;
                default:
                    o = value;
                    break;
            }

            return o;
        }
    }
}