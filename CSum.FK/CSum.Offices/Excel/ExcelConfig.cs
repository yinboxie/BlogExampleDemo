using System.Collections.Generic;
using System.Drawing;

namespace CSum.Offices.Excel
{
    /// <summary>
    ///     描 述：Excel导入导出设置
    /// </summary>
    public class ExcelConfig
    {
        private string _headfont;
        private short _headpoint;
        private string _titlefont;
        private short _titlepoint;

        /// <summary>
        ///     文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        ///     标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     前景色
        /// </summary>
        public Color ForeColor { get; set; }

        /// <summary>
        ///     背景色
        /// </summary>
        public Color Background { get; set; }

        /// <summary>
        ///     标题字号
        /// </summary>
        public short TitlePoint
        {
            get
            {
                if (_titlepoint == 0)
                    return 20;
                return _titlepoint;
            }
            set => _titlepoint = value;
        }

        /// <summary>
        ///     列头字号
        /// </summary>
        public short HeadPoint
        {
            get
            {
                if (_headpoint == 0)
                    return 10;
                return _headpoint;
            }
            set => _headpoint = value;
        }

        /// <summary>
        ///     标题高度
        /// </summary>
        public short TitleHeight { get; set; }

        /// <summary>
        ///     列标题高度
        /// </summary>
        public short HeadHeight { get; set; }

        /// <summary>
        ///     标题字体
        /// </summary>
        public string TitleFont
        {
            get
            {
                if (_titlefont == null)
                    return "微软雅黑";
                return _titlefont;
            }
            set => _titlefont = value;
        }

        /// <summary>
        ///     列头字体
        /// </summary>
        public string HeadFont
        {
            get
            {
                if (_headfont == null)
                    return "微软雅黑";
                return _headfont;
            }
            set => _headfont = value;
        }

        /// <summary>
        ///     是否按内容长度来适应表格宽度
        /// </summary>
        public bool IsAllSizeColumn { get; set; }

        /// <summary>
        ///     列设置
        /// </summary>
        public List<ColumnEntity> ColumnEntity { get; set; }
    }
}