using System;
using System.IO;
using System.Web;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CSum.Offices.Excel
{
    public class ExcelTemplate
    {
        private readonly string newFileName;
        private readonly string templatePath;
        private readonly string templdateName;

        public ExcelTemplate(string templdateName, string newFileName)
        {
            SheetName = "sheet1";
            templatePath = HttpContext.Current.Server.MapPath("/") + "/Resource/ExcelTemplate/";
            this.templdateName = string.Format("{0}{1}", templatePath, templdateName);
            this.newFileName = newFileName;
        }

        public string SheetName { get; set; }

        public void ExportDataToExcel(Action<ISheet> actionMethod)
        {
            using (var ms = SetDataToExcel(actionMethod))
            {
                var data = ms.ToArray();

                #region response to the client

                var response = HttpContext.Current.Response;
                response.Clear();
                response.Charset = "UTF-8";
                response.ContentType = "application/vnd-excel"; //"application/vnd.ms-excel";
                HttpContext.Current.Response.AddHeader("Content-Disposition",
                    string.Format("attachment; filename=" + newFileName));
                HttpContext.Current.Response.BinaryWrite(data);

                #endregion
            }
        }

        private MemoryStream SetDataToExcel(Action<ISheet> actionMethod)
        {
            //Load template file
            var file = new FileStream(templdateName, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(file);
            var sheet = workbook.GetSheet(SheetName);

            if (actionMethod != null) actionMethod(sheet);

            sheet.ForceFormulaRecalculation = true;
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                //ms.Position = 0;
                return ms;
            }
        }
    }
}