using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSum.Model
{
    public class ExportModel
    {
        /// <summary>
        /// 保存的文件名称
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 表头集合
        /// </summary>
        public List<ExportDataColumn> LstCol { get; set; }
        /// <summary>
        /// 数据
        /// </summary>
        public string Data { get; set; }
    }
    /// <summary>
    /// 列的list
    /// </summary>
    public class ExportDataColumn
    {
        public string prop { get; set; }
        public string label { get; set; }
        public int width { get; set; }
    }
}
