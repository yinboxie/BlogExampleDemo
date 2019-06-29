let props = {
  // 远程请求路径
  url: {
    type: String
  },
  // 查询参数对象，无论是单条件查询，高级查询，远程请求数据或本地请求数据，都是改变该参数
  params: {
    type: [Object, Array]
  },
  // 自动加载,组件初始化时就加载
  autoLoad: {
    type: Boolean,
    default: true
  },
  // 每页数量
  pageSize: {
    type: Number,
    default: 10
  },
  // 排序字段
  sortName: {
    type: String,
    default: 'CreateDate'
  },
  // 排序顺序
  sordName: {
    type: String,
    default: 'desc',
    validator: value => {
      const sordTypes = ['desc', 'asc']
      return sordTypes.indexOf(value.toLowerCase()) !== -1
    }
  },

  /** 以下字段用来自定义分页或排序字段 */
  // 远程请求，返回的数据字段名称
  listField: {
    type: String,
    default: 'rows'
  },
  // 远程请求，返回的总数字段名称
  totalField: {
    type: String,
    default: 'total'
  },
  // 远程请求 请求的参数数据 当前页字段名称
  pageIndexField: {
    type: String,
    default: 'page'
  },
  // 远程请求 请求的参数数据 每页行数字段名称
  pageSizeField: {
    type: String,
    default: 'rows'
  },
  // 远程请求 请求的参数数据 排序字段名称
  sortNameField: {
    type: String,
    default: 'sidx'
  },
  // 远程请求 请求的参数数据 排序顺序字段名称
  sordField: {
    type: String,
    default: 'sord'
  }
}
export default props
