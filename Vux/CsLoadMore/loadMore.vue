<template>
  <div>
    <div v-for="(item,idx) in tableData"
         :key="idx"
         class="box">
      <slot :item="item"></slot>
    </div>
    <load-more v-if="loading"
               tip="正在加载"></load-more>
    <load-more v-else
               :show-loading="false"
               @click.prevent.native="load"
               :tip="tipText"
               background-color="#fbf9fe"></load-more>
  </div>
</template>
<script>
import props from './props'
import { LoadMore } from 'vux'
import axios from '@/http/config'
export default {
  name: 'CsLoadMore',
  props,
  data () {
    const _this = this
    return {
      tableData: [], // 列表数据
      loading: false,
      isLoadMore: true,
      // 查询参数
      queryJson: (() => {
        const { params } = _this
        return params
      })(),

      pageIndex: 1, // 当前页
      total: 0 // 数据总条数
    }
  },
  computed: {
    tipText () {
      // 暂无数据, 没有更多数据, 轻按加载更多
      if (!this.tableData || this.tableData.length === 0) {
        return '暂无数据'
      }
      return this.isLoadMore ? '轻按加载更多' : '没有更多数据'
    }
  },
  methods: {
    load () {
      if (!this.isLoadMore) {
        return
      }
      this.fetch()
    },
    fetch () {
      this.loading = true
      let { url, pageSize, pageIndex, sortName, sordName, listField, totalField,
        pageIndexField, pageSizeField, sortNameField, sordField } = this

      let params = Object.assign({}, this.queryJson)

      // 分页参数
      params = Object.assign(params, {
        [pageIndexField]: pageIndex,
        [pageSizeField]: pageSize
      })
      // 排序参数
      params = Object.assign(params, {
        [sortNameField]: sortName,
        [sordField]: sordName
      })
      axios.get(url, { params }).then(response => {
        this.total = response[totalField] // 总数
        let result = response[listField] // 当次加载的数据

        // 是否还可以加载更多 此种逻辑设计存在缺陷,如果在浏览列表的同时,又增加了新的记录
        this.isLoadMore = result.length === pageSize
        this.pageIndex++
        for (let item of result) {
          this.tableData.push(item)
        }
      }).catch(error => {
        console.error('获取数据失败 ', error)
      }).finally(() => {
        this.loading = false
      })
    }
  },
  mounted () {
    if (this.autoLoad) {
      this.fetch()
    }
  },
  watch: {
    params: function (val) {
      this.queryJson = val
      this.pageIndex = 1
      this.tableData = []
      this.fetch()
    }
  },
  components: {
    LoadMore
  }
}
</script>
<style lang="scss" scoped>
.box {
  border-bottom: 1px solid #d9d9d9;
}
</style>
