using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using NPOI.HPSF;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace CSum.Offices.Excel
{
    /// <summary>
    ///     描 述：NPOI Excel DataTable操作类
    public class ExcelHelper
    {
        #region DataTable导出到Excel文件excelConfig中FileName设置为全路径

        /// <summary>
        ///     DataTable导出到Excel文件 Export()
        /// </summary>
        /// <param name="dtSource">DataTable数据源</param>
        /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
        public static void ExcelExport(DataTable dtSource, ExcelConfig excelConfig)
        {
            using (var ms = ExportMemoryStream(dtSource, excelConfig))
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
        public static MemoryStream ExportMemoryStream(DataTable dtSource, ExcelConfig excelConfig)
        {
            /*var dt = dtSource.Clone();
            int len = dt.Columns.Count;
            for (int i = 0; i < len;i++)
            {
                DataColumn column = dt.Columns[i];
                if (excelConfig.ColumnEntity.Find(p => p.Column == column.ColumnName) == null)
                {
                    dtSource.Columns.Remove(column.ColumnName);
                }
            }*/
            for (var i = 0; i < excelConfig.ColumnEntity.Count; i++)
            {
                var modelConfig = excelConfig.ColumnEntity[i];
                if (!dtSource.Columns.Contains(modelConfig.Column)) //如果源数据不包含该列
                    dtSource.Columns.Add(modelConfig.Column, typeof(string), "");
            }

            var arrCol = excelConfig.ColumnEntity.Select(p => p.Column).ToArray();
            dtSource = dtSource.DefaultView.ToTable(true, arrCol);
            //筛选出配置和数据源都存在的列
            var eColumns = new List<ColumnEntity>();
            foreach (var item in excelConfig.ColumnEntity)
                if (dtSource.Columns[item.Column] != null)
                    eColumns.Add(item);

            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet();

            #region 右击文件 属性信息

            {
                var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "NPOI";
                workbook.DocumentSummaryInformation = dsi;

                var si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "system"; //填加xls文件作者信息
                si.ApplicationName = "中夏信科"; //填加xls文件创建程序信息
                si.LastAuthor = "system"; //填加xls文件最后保存者信息
                si.Comments = "system"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "主题信息"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion

            #region 设置标题样式

            var headStyle = workbook.CreateCellStyle();
            var arrColWidth = new int[dtSource.Columns.Count];
            var arrColName = new string[dtSource.Columns.Count]; //列名
            var arryColumStyle = new ICellStyle[dtSource.Columns.Count]; //样式表
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

            foreach (var eColumn in eColumns)
            {
                var item = dtSource.Columns[eColumn.Column];
                var columnStyle = workbook.CreateCellStyle();
                columnStyle.Alignment = HorizontalAlignment.Center;
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length;
                arrColName[item.Ordinal] = item.ColumnName;
                if (excelConfig.ColumnEntity != null)
                {
                    var columnentity = excelConfig.ColumnEntity.Find(t => t.Column == item.ColumnName);
                    if (columnentity != null)
                    {
                        arrColName[item.Ordinal] = columnentity.ExcelColumn;
                        if (columnentity.Width != 0) arrColWidth[item.Ordinal] = columnentity.Width;
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

                        columnStyle.Alignment = getAlignment(columnentity.Alignment);
                    }
                }

                arryColumStyle[item.Ordinal] = columnStyle;
            }

            //foreach (DataColumn item in dtSource.Columns)
            //{
            //    ICellStyle columnStyle = workbook.CreateCellStyle();
            //    columnStyle.Alignment = HorizontalAlignment.Center;
            //    arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            //    arrColName[item.Ordinal] = item.ColumnName.ToString();
            //    if (excelConfig.ColumnEntity != null)
            //    {
            //        ColumnEntity columnentity = excelConfig.ColumnEntity.Find(t => t.Column == item.ColumnName);
            //        if (columnentity != null)
            //        {
            //            arrColName[item.Ordinal] = columnentity.ExcelColumn;
            //            if (columnentity.Width != 0)
            //            {
            //                arrColWidth[item.Ordinal] = columnentity.Width;
            //            }
            //            if (columnentity.Background != new Color())
            //            {
            //                if (columnentity.Background != new Color())
            //                {
            //                    columnStyle.FillPattern = FillPattern.SolidForeground;
            //                    columnStyle.FillForegroundColor = GetXLColour(workbook, columnentity.Background);
            //                }
            //            }
            //            if (columnentity.Font != null || columnentity.Point != 0 || columnentity.ForeColor != new Color())
            //            {
            //                IFont columnFont = workbook.CreateFont();
            //                columnFont.FontHeightInPoints = 10;
            //                if (columnentity.Font != null)
            //                {
            //                    columnFont.FontName = columnentity.Font;
            //                }
            //                if (columnentity.Point != 0)
            //                {
            //                    columnFont.FontHeightInPoints = columnentity.Point;
            //                }
            //                if (columnentity.ForeColor != new Color())
            //                {
            //                    columnFont.Color = GetXLColour(workbook, columnentity.ForeColor);
            //                }
            //                columnStyle.SetFont(font);
            //            }
            //            columnStyle.Alignment = getAlignment(columnentity.Alignment);
            //        }
            //    }
            //    arryColumStyle[item.Ordinal] = columnStyle;
            //}
            if (excelConfig.IsAllSizeColumn)
            {
                #region 根据列中最长列的长度取得列宽

                for (var i = 0; i < dtSource.Rows.Count; i++)
                for (var j = 0; j < dtSource.Columns.Count; j++)
                    if (arrColWidth[j] != 0)
                    {
                        var intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                        if (intTemp > arrColWidth[j]) arrColWidth[j] = intTemp;
                    }

                #endregion
            }

            #endregion

            var rowIndex = 0;

            #region 表头及样式

            if (excelConfig.Title != null)
            {
                var headerRow = sheet.CreateRow(rowIndex);
                rowIndex++;
                if (excelConfig.TitleHeight != 0) headerRow.Height = (short) (excelConfig.TitleHeight * 20);
                headerRow.HeightInPoints = 25;
                headerRow.CreateCell(0).SetCellValue(excelConfig.Title);
                headerRow.GetCell(0).CellStyle = headStyle;
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1)); // ------------------
            }

            #endregion

            #region 列头及样式

            {
                var headerRow = sheet.CreateRow(rowIndex);
                rowIndex++;

                #region 如果设置了列标题就按列标题定义列头，没定义直接按字段名输出

                for (var i = 0; i < eColumns.Count; i++)
                {
                    var eColumn = eColumns[i];
                    var column = dtSource.Columns[eColumn.Column];
                    headerRow.CreateCell(i).SetCellValue(arrColName[column.Ordinal]);
                    headerRow.GetCell(i).CellStyle = cHeadStyle;
                    //设置列宽
                    sheet.SetColumnWidth(i, (arrColWidth[column.Ordinal] + 1) * 256);
                }

                //foreach (DataColumn column in dtSource.Columns)
                //{
                //    headerRow.CreateCell(column.Ordinal).SetCellValue(arrColName[column.Ordinal]);
                //    headerRow.GetCell(column.Ordinal).CellStyle = cHeadStyle;
                //    //设置列宽
                //    sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                //}

                #endregion
            }

            #endregion

            var dateStyle = workbook.CreateCellStyle();
            var format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535)
                {
                    sheet = workbook.CreateSheet();
                    rowIndex = 0;

                    #region 表头及样式

                    {
                        if (excelConfig.Title != null)
                        {
                            var headerRow = sheet.CreateRow(rowIndex);
                            rowIndex++;
                            if (excelConfig.TitleHeight != 0) headerRow.Height = (short) (excelConfig.TitleHeight * 20);
                            headerRow.HeightInPoints = 25;
                            headerRow.CreateCell(0).SetCellValue(excelConfig.Title);
                            headerRow.GetCell(0).CellStyle = headStyle;
                            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0,
                                dtSource.Columns.Count - 1)); // ------------------
                        }
                    }

                    #endregion

                    #region 列头及样式

                    {
                        var headerRow = sheet.CreateRow(rowIndex);
                        rowIndex++;

                        #region 如果设置了列标题就按列标题定义列头，没定义直接按字段名输出

                        for (var i = 0; i < eColumns.Count; i++)
                        {
                            var eColumn = eColumns[i];
                            var column = dtSource.Columns[eColumn.Column];
                            headerRow.CreateCell(i).SetCellValue(arrColName[column.Ordinal]);
                            headerRow.GetCell(i).CellStyle = cHeadStyle;
                            //设置列宽
                            sheet.SetColumnWidth(i, (arrColWidth[column.Ordinal] + 1) * 256);
                        }

                        //foreach (DataColumn column in dtSource.Columns)
                        //{
                        //    headerRow.CreateCell(column.Ordinal).SetCellValue(arrColName[column.Ordinal]);
                        //    headerRow.GetCell(column.Ordinal).CellStyle = cHeadStyle;
                        //    //设置列宽
                        //    sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                        //}

                        #endregion
                    }

                    #endregion
                }

                #endregion

                #region 填充内容

                var dataRow = sheet.CreateRow(rowIndex);
                for (var i = 0; i < eColumns.Count; i++)
                {
                    var eColumn = eColumns[i];
                    var column = dtSource.Columns[eColumn.Column];
                    var newCell = dataRow.CreateCell(i);
                    newCell.CellStyle = arryColumStyle[column.Ordinal];
                    var drValue = row[column] == null ? "" : row[column].ToString();
                    ;
                    SetCell(newCell, dateStyle, column.DataType, drValue);
                }

                //foreach (DataColumn column in dtSource.Columns)
                //{
                //    ICell newCell = dataRow.CreateCell(column.Ordinal);
                //    newCell.CellStyle = arryColumStyle[column.Ordinal];
                //    string drValue = row[column].ToString();
                //    SetCell(newCell, dateStyle, column.DataType, drValue);
                //}

                #endregion

                rowIndex++;
            }

            //锁定标题行和表头行
            sheet.CreateFreezePane(0, 2);

            var ms = new MemoryStream();
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            return ms;
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

        #region 设置列的对齐方式

        /// <summary>
        ///     设置对齐方式
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        private static HorizontalAlignment getAlignment(string style)
        {
            switch (style)
            {
                case "center":
                    return HorizontalAlignment.Center;
                case "left":
                    return HorizontalAlignment.Left;
                case "right":
                    return HorizontalAlignment.Right;
                case "fill":
                    return HorizontalAlignment.Fill;
                case "justify":
                    return HorizontalAlignment.Justify;
                case "centerselection":
                    return HorizontalAlignment.CenterSelection;
                case "distributed":
                    return HorizontalAlignment.Distributed;
            }

            return HorizontalAlignment.General;
        }

        #endregion

        #region Excel导出方法 ExcelDownload

        /// <summary>
        ///     Excel导出下载
        /// </summary>
        /// <param name="dtSource">DataTable数据源</param>
        /// <param name="excelConfig">导出设置包含文件名、标题、列设置</param>
        public static void ExcelDownload(DataTable dtSource, ExcelConfig excelConfig)
        {
            var curContext = HttpContext.Current;
            // 设置编码和附件格式
            curContext.Response.ContentType = "application/ms-excel";
            curContext.Response.ContentEncoding = Encoding.UTF8;
            curContext.Response.Charset = "";
            curContext.Response.AppendHeader("Content-Disposition",
                "attachment;filename=" + HttpUtility.UrlEncode(excelConfig.FileName, Encoding.UTF8));
            //调用导出具体方法Export()
            curContext.Response.BinaryWrite(ExportMemoryStream(dtSource, excelConfig).GetBuffer());
            curContext.Response.End();
        }

        /// <summary>
        ///     Excel导出下载
        /// </summary>
        /// <param name="list">数据源</param>
        /// <param name="templdateName">模板文件名</param>
        /// <param name="newFileName">文件名</param>
        public static void ExcelDownload(List<TemplateMode> list, string templdateName, string newFileName)
        {
            var response = HttpContext.Current.Response;
            response.Clear();
            response.Charset = "UTF-8";
            response.ContentType = "application/vnd-excel"; //"application/vnd.ms-excel";
            HttpContext.Current.Response.AddHeader("Content-Disposition",
                string.Format("attachment; filename=" + newFileName));
            HttpContext.Current.Response.BinaryWrite(ExportListByTempale(list, templdateName).ToArray());
        }

        #endregion

        #region ListExcel导出(加载模板)

        /// <summary>
        ///     List根据模板导出ExcelMemoryStream
        /// </summary>
        /// <param name="list"></param>
        /// <param name="templdateName"></param>
        public static MemoryStream ExportListByTempale(List<TemplateMode> list, string templdateName)
        {
            var templatePath = HttpContext.Current.Server.MapPath("/") + "/Resource/ExcelTemplate/";
            var templdateName1 = string.Format("{0}{1}", templatePath, templdateName);

            var fileStream = new FileStream(templdateName1, FileMode.Open, FileAccess.Read);
            ISheet sheet = null;
            if (templdateName.IndexOf(".xlsx") == -1) //2003
            {
                var hssfworkbook = new HSSFWorkbook(fileStream);
                sheet = hssfworkbook.GetSheetAt(0);
                SetPurchaseOrder(sheet, list);
                sheet.ForceFormulaRecalculation = true;
                using (var ms = new MemoryStream())
                {
                    hssfworkbook.Write(ms);
                    ms.Flush();
                    return ms;
                }
            }

            var xssfworkbook = new XSSFWorkbook(fileStream);
            sheet = xssfworkbook.GetSheetAt(0);
            SetPurchaseOrder(sheet, list);
            sheet.ForceFormulaRecalculation = true;
            using (var ms = new MemoryStream())
            {
                xssfworkbook.Write(ms);
                ms.Flush();
                return ms;
            }
        }

        /// <summary>
        ///     赋值单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="list"></param>
        private static void SetPurchaseOrder(ISheet sheet, List<TemplateMode> list)
        {
            foreach (var item in list)
            {
                IRow row = null;
                ICell cell = null;
                row = sheet.GetRow(item.row);
                if (row == null) row = sheet.CreateRow(item.row);
                cell = row.GetCell(item.cell);
                if (cell == null) cell = row.CreateCell(item.cell);
                cell.SetCellValue(item.value);
            }
        }

        #endregion

        #region 从Excel导入

        /// <summary>
        ///     读取excel ,默认第一行为标头
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public static DataTable ExcelImport(string strFileName)
        {
            var dt = new DataTable();

            ISheet sheet = null;
            using (var file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                if (strFileName.IndexOf(".xlsx") == -1) //2003
                {
                    var hssfworkbook = new HSSFWorkbook(file);
                    sheet = hssfworkbook.GetSheetAt(0);
                }
                else //2007
                {
                    var xssfworkbook = new XSSFWorkbook(file);
                    sheet = xssfworkbook.GetSheetAt(0);
                }
            }

            var rows = sheet.GetRowEnumerator();

            var headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            for (var j = 0; j < cellCount; j++)
            {
                var cell = headerRow.GetCell(j);
                dt.Columns.Add(cell.ToString());
            }

            for (var i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var dataRow = dt.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                    if (row.GetCell(j) != null)
                        dataRow[j] = row.GetCell(j).ToString();

                dt.Rows.Add(dataRow);
            }

            return dt;
        }

        /// <summary>
        ///     读取excel ,默认第一行为标头
        /// </summary>
        /// <param name="fileStream">文件数据流</param>
        /// <returns></returns>
        public static DataTable ExcelImport(Stream fileStream, string flieType)
        {
            var dt = new DataTable();
            ISheet sheet = null;
            if (flieType == ".xls")
            {
                var hssfworkbook = new HSSFWorkbook(fileStream);
                sheet = hssfworkbook.GetSheetAt(0);
            }
            else
            {
                var xssfworkbook = new XSSFWorkbook(fileStream);
                sheet = xssfworkbook.GetSheetAt(0);
            }

            var rows = sheet.GetRowEnumerator();
            var headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            for (var j = 0; j < cellCount; j++)
            {
                var cell = headerRow.GetCell(j);
                dt.Columns.Add(cell.ToString());
            }

            for (var i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var dataRow = dt.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                    if (row.GetCell(j) != null)
                        dataRow[j] = row.GetCell(j).ToString();

                dt.Rows.Add(dataRow);
            }

            return dt;
        }

        /// <summary>
        ///     读取excel ,默认第一行为标头
        ///     读取所有页
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public static DataTable ExcelImport(string strFileName, List<string> lstCol, bool allSheets)
        {
            if (!allSheets) return ExcelImport(strFileName);
            var dt = DataTableHelper.CreateTable(lstCol);

            ISheet sheet = null;
            var countSheet = 0;
            using (var file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                if (strFileName.IndexOf(".xlsx") == -1) //2003
                {
                    var hssfworkbook = new HSSFWorkbook(file);
                    countSheet = hssfworkbook.NumberOfSheets;
                    // sheet = hssfworkbook.GetSheetAt(0);
                }
                else //2007
                {
                    var xssfworkbook = new XSSFWorkbook(file);
                    countSheet = xssfworkbook.NumberOfSheets;
                    //sheet = xssfworkbook.GetSheetAt(0);
                }
            }

            for (var i = 0; i < countSheet; i++) DataTableHelper.AddTableData(dt, ExcelImport(strFileName, i, lstCol));
            return dt;
        }

        /// <summary>
        ///     读取excel ,默认第一行为标头
        ///     读取所有页
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        private static DataTable ExcelImport(string strFileName, int sheetsIndex, List<string> lstCol)
        {
            var dt = DataTableHelper.CreateTable(lstCol);

            ISheet sheet = null;
            using (var file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                if (strFileName.IndexOf(".xlsx") == -1) //2003
                {
                    var hssfworkbook = new HSSFWorkbook(file);
                    sheet = hssfworkbook.GetSheetAt(sheetsIndex);
                }
                else //2007
                {
                    var xssfworkbook = new XSSFWorkbook(file);
                    sheet = xssfworkbook.GetSheetAt(sheetsIndex);
                }
            }

            var rows = sheet.GetRowEnumerator();

            var headerRow = sheet.GetRow(0);
            if (headerRow == null) return dt;
            int cellCount = headerRow.LastCellNum;
            var lstCell = new List<CellDto>();
            for (var j = 0; j < cellCount; j++)
            {
                var cell = headerRow.GetCell(j);
                if (cell == null) continue;
                if (lstCol.Contains(cell.ToString())) lstCell.Add(new CellDto {Name = cell.ToString(), Index = j});
            }


            for (var i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var dataRow = dt.NewRow();

                //for (int j = row.FirstCellNum; j < cellCount; j++)
                //{
                //    if (row.GetCell(j) != null)
                //        dataRow[j] = row.GetCell(j).ToString();
                //}
                foreach (var item in lstCell)
                {
                    var cell = row.GetCell(item.Index);
                    if (cell != null)
                        dataRow[item.Name] = cell.ToString();
                }

                dt.Rows.Add(dataRow);
            }

            return dt;
        }

        #endregion
    }

    public class CellDto
    {
        /// <summary>
        ///     列名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        ///     序号
        /// </summary>
        public int Index { get; set; }
    }
}