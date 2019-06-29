using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace CSum.Offices
{
    /// <summary>
    ///     DataTable操作辅助类
    /// </summary>
    public class DataTableHelper
    {
        public static void AddTableData(DataTable dt1, DataTable dt2)
        {
            if (dt2 == null || dt2.Rows.Count == 0) return;
            // 改进后的方法
            DataRow drcalc;
            foreach (DataRow dr in dt2.Rows)
            {
                drcalc = dt1.NewRow();
                drcalc.ItemArray = dr.ItemArray;
                dt1.Rows.Add(drcalc);
            }
        }

        /// <summary>
        ///     给DataTable增加一个自增列
        ///     如果DataTable 存在 identityid 字段  则 直接返回DataTable 不做任何处理
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <returns>返回Datatable 增加字段 identityid </returns>
        public static DataTable AddIdentityColumn(DataTable dt)
        {
            if (!dt.Columns.Contains("identityid"))
            {
                dt.Columns.Add("identityid");
                for (var i = 0; i < dt.Rows.Count; i++) dt.Rows[i]["identityid"] = (i + 1).ToString();
            }

            return dt;
        }

        /// <summary>
        ///     检查DataTable 是否有数据行
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <returns></returns>
        public static bool IsHaveRows(DataTable dt)
        {
            if (dt != null && dt.Rows.Count > 0)
                return true;

            return false;
        }

        /// <summary>
        ///     DataTable转换成实体列表
        /// </summary>
        /// <typeparam name="T">实体 T </typeparam>
        /// <param name="table">datatable</param>
        /// <returns></returns>
        public static IList<T> DataTableToList<T>(DataTable table)
            where T : class
        {
            if (!IsHaveRows(table))
                return new List<T>();

            IList<T> list = new List<T>();
            var model = default(T);
            foreach (DataRow dr in table.Rows)
            {
                model = Activator.CreateInstance<T>();

                foreach (DataColumn dc in dr.Table.Columns)
                {
                    var drValue = dr[dc.ColumnName];
                    var pi = model.GetType().GetProperty(dc.ColumnName);

                    if (pi != null && pi.CanWrite && drValue != null && !Convert.IsDBNull(drValue))
                        pi.SetValue(model, drValue, null);
                }

                list.Add(model);
            }

            return list;
        }

        /// <summary>
        ///     实体列表转换成DataTable
        /// </summary>
        /// <typeparam name="T">实体</typeparam>
        /// <param name="list"> 实体列表</param>
        /// <returns></returns>
        public static DataTable ListToDataTable<T>(IList<T> list)
            where T : class
        {
            if (list == null || list.Count <= 0) return null;
            var dt = new DataTable(typeof(T).Name);
            DataColumn column;
            DataRow row;

            var myPropertyInfo = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var length = myPropertyInfo.Length;
            var createColumn = true;

            foreach (var t in list)
            {
                if (t == null) continue;

                row = dt.NewRow();
                for (var i = 0; i < length; i++)
                {
                    var pi = myPropertyInfo[i];
                    var name = pi.Name;
                    if (createColumn)
                    {
                        column = new DataColumn(name, pi.PropertyType);
                        dt.Columns.Add(column);
                    }

                    row[name] = pi.GetValue(t, null);
                }

                if (createColumn) createColumn = false;

                dt.Rows.Add(row);
            }

            return dt;
        }

        /// <summary>
        ///     将泛型集合类转换成DataTable
        /// </summary>
        /// <typeparam name="T">集合项类型</typeparam>
        /// <param name="list">集合</param>
        /// <returns>数据集(表)</returns>
        public static DataTable ToDataTable<T>(IList<T> list)
        {
            return ToDataTable(list, null);
        }

        /// <summary>
        ///     将泛型集合类转换成DataTable
        /// </summary>
        /// <typeparam name="T">集合项类型</typeparam>
        /// <param name="list">集合</param>
        /// <param name="propertyName">需要返回的列的列名</param>
        /// <returns>数据集(表)</returns>
        public static DataTable ToDataTable<T>(IList<T> list, params string[] propertyName)
        {
            var propertyNameList = new List<string>();
            if (propertyName != null)
                propertyNameList.AddRange(propertyName);

            var result = new DataTable();
            if (list.Count > 0)
            {
                var propertys = list[0].GetType().GetProperties();
                foreach (var pi in propertys)
                    if (propertyNameList.Count == 0)
                    {
                        result.Columns.Add(pi.Name, pi.PropertyType);
                    }
                    else
                    {
                        if (propertyNameList.Contains(pi.Name)) result.Columns.Add(pi.Name, pi.PropertyType);
                    }

                for (var i = 0; i < list.Count; i++)
                {
                    var tempList = new ArrayList();
                    foreach (var pi in propertys)
                        if (propertyNameList.Count == 0)
                        {
                            var obj = pi.GetValue(list[i], null);
                            tempList.Add(obj);
                        }
                        else
                        {
                            if (propertyNameList.Contains(pi.Name))
                            {
                                var obj = pi.GetValue(list[i], null);
                                tempList.Add(obj);
                            }
                        }

                    var array = tempList.ToArray();
                    result.LoadDataRow(array, true);
                }
            }

            return result;
        }

        /// <summary>
        ///     根据nameList里面的字段创建一个表格,返回该表格的DataTable
        /// </summary>
        /// <param name="nameList">包含字段信息的列表</param>
        /// <returns>DataTable</returns>
        public static DataTable CreateTable(List<string> nameList)
        {
            if (nameList.Count <= 0)
                return null;

            var myDataTable = new DataTable();
            foreach (var columnName in nameList) myDataTable.Columns.Add(columnName, typeof(string));
            return myDataTable;
        }

        /// <summary>
        ///     通过字符列表创建表字段，字段格式可以是：
        ///     1) a,b,c,d,e
        ///     2) a|int,b|string,c|bool,d|decimal
        /// </summary>
        /// <param name="nameString"></param>
        /// <returns></returns>
        public static DataTable CreateTable(string nameString)
        {
            var nameArray = nameString.Split(',', ';');
            var nameList = new List<string>();
            var dt = new DataTable();
            foreach (var item in nameArray)
                if (!string.IsNullOrEmpty(item))
                {
                    var subItems = item.Split('|');
                    if (subItems.Length == 2)
                        dt.Columns.Add(subItems[0], ConvertType(subItems[1]));
                    else
                        dt.Columns.Add(subItems[0]);
                }

            return dt;
        }

        private static Type ConvertType(string typeName)
        {
            typeName = typeName.ToLower().Replace("system.", "");
            var newType = typeof(string);
            switch (typeName)
            {
                case "boolean":
                case "bool":
                    newType = typeof(bool);
                    break;
                case "int16":
                case "short":
                    newType = typeof(short);
                    break;
                case "int32":
                case "int":
                    newType = typeof(int);
                    break;
                case "long":
                case "int64":
                    newType = typeof(long);
                    break;
                case "uint16":
                case "ushort":
                    newType = typeof(ushort);
                    break;
                case "uint32":
                case "uint":
                    newType = typeof(uint);
                    break;
                case "uint64":
                case "ulong":
                    newType = typeof(ulong);
                    break;
                case "single":
                case "float":
                    newType = typeof(float);
                    break;

                case "string":
                    newType = typeof(string);
                    break;
                case "guid":
                    newType = typeof(Guid);
                    break;
                case "decimal":
                    newType = typeof(decimal);
                    break;
                case "double":
                    newType = typeof(double);
                    break;
                case "datetime":
                    newType = typeof(DateTime);
                    break;
                case "byte":
                    newType = typeof(byte);
                    break;
                case "char":
                    newType = typeof(char);
                    break;
            }

            return newType;
        }

        /// <summary>
        ///     获得从DataRowCollection转换成的DataRow数组
        /// </summary>
        /// <param name="drc">DataRowCollection</param>
        /// <returns></returns>
        public static DataRow[] GetDataRowArray(DataRowCollection drc)
        {
            var count = drc.Count;
            var drs = new DataRow[count];
            for (var i = 0; i < count; i++) drs[i] = drc[i];
            return drs;
        }

        /// <summary>
        ///     将DataRow数组转换成DataTable，注意行数组的每个元素须具有相同的数据结构，
        ///     否则当有元素长度大于第一个元素时，抛出异常
        /// </summary>
        /// <param name="rows">行数组</param>
        /// <returns></returns>
        public static DataTable GetTableFromRows(DataRow[] rows)
        {
            if (rows.Length <= 0) return new DataTable();
            var dt = rows[0].Table.Clone();
            dt.DefaultView.Sort = rows[0].Table.DefaultView.Sort;
            for (var i = 0; i < rows.Length; i++) dt.LoadDataRow(rows[i].ItemArray, true);
            return dt;
        }

        /// <summary>
        ///     排序表的视图
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sorts"></param>
        /// <returns></returns>
        public static DataTable SortedTable(DataTable dt, params string[] sorts)
        {
            if (dt.Rows.Count > 0)
            {
                var tmp = "";
                for (var i = 0; i < sorts.Length; i++) tmp += sorts[i] + ",";
                dt.DefaultView.Sort = tmp.TrimEnd(',');
            }

            return dt;
        }

        /// <summary>
        ///     根据条件过滤表的内容
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        public static DataTable FilterDataTable(DataTable dt, string condition)
        {
            if (condition.Trim() == "") return dt;
            var newdt = new DataTable();
            newdt = dt.Clone();
            var dr = dt.Select(condition);
            for (var i = 0; i < dr.Length; i++) newdt.ImportRow(dr[i]);
            return newdt;
        }

        /// <summary>
        ///     转换.NET的Type到数据库参数的类型
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        public static DbType TypeToDbType(Type t)
        {
            DbType dbt;
            try
            {
                dbt = (DbType) Enum.Parse(typeof(DbType), t.Name);
            }
            catch
            {
                dbt = DbType.Object;
            }

            return dbt;
        }

        /// <summary>
        ///     使用分隔符串联表格字段的内容,如：a,b,c
        /// </summary>
        /// <param name="dt">表格</param>
        /// <param name="columnName">字段名称</param>
        /// <param name="append">增加的字符串，无则为空</param>
        /// <param name="splitChar">分隔符，如逗号(,)</param>
        /// <returns></returns>
        public static string ConcatColumnValue(DataTable dt, string columnName, string append, char splitChar)
        {
            var result = append;
            if (dt != null && dt.Rows.Count > 0)
                foreach (DataRow row in dt.Rows)
                    result += string.Format("{0}{1}", splitChar, row[columnName]);
            return result.Trim(splitChar);
        }

        /// <summary>
        ///     使用逗号串联表格字段的内容,如：a,b,c
        /// </summary>
        /// <param name="dt">表格</param>
        /// <param name="columnName">字段名称</param>
        /// <param name="append">增加的字符串，无则为空</param>
        /// <returns></returns>
        public static string ConcatColumnValue(DataTable dt, string columnName, string append)
        {
            var result = append;
            if (dt != null && dt.Rows.Count > 0)
                foreach (DataRow row in dt.Rows)
                    result += string.Format(",{0}", row[columnName]);
            return result.Trim(',');
        }
    }
}