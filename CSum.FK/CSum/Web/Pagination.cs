namespace CSum
{
    /// <summary>
    ///     分页参数
    /// </summary>
    public class Pagination
    {
        public Pagination()
        {
            sidx = "";
            sord = "";
        }

        /// <summary>
        ///     每页行数
        /// </summary>
        public int rows { get; set; }

        /// <summary>
        ///     当前页(当值为-1时查询所有数据)
        /// </summary>
        public int page { get; set; } = -1;

        /// <summary>
        ///     排序列
        /// </summary>
        public string sidx { get; set; }

        /// <summary>
        ///     排序类型
        /// </summary>
        public string sord { get; set; }

        /// <summary>
        ///     总记录数
        /// </summary>
        public int records { get; set; }

        /// <summary>
        ///     总页数
        /// </summary>
        public int total
        {
            get
            {
                if (records > 0)
                    return records % rows == 0 ? records / rows : records / rows + 1;
                return 0;
            }
        }

        ///// <summary>
        ///// 查询条件Json
        ///// </summary>
        //public string conditionJson { get; set; }
    }
}