interface IFilter {
  type: 'rule'
  filterBy: {
    type: 'string' | 'number' | 'boolean'
    value: string | number | boolean
  },
  filterValue: string | number,
  filterOperator: 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le'
}
type JoinOption = 'and' | 'or' | 'not'

interface IFilterGroup {
  type: 'group'
  items: (IFilterGroup | IFilter)[]
  joinedBy?: JoinOption
}
interface IFilterByOptions {
  value: string,
  type: 'string' | 'boolean' | 'number'
}
function buildFilterQuery(groups: IFilterGroup[], joinedBy: JoinOption) {
  if (groups.length < 1) {
    return ""
  }
  const groupFilters = groups.map(processFilters)
  if (groupFilters.length === 0) return "";
  const filterStrings = groupFilters.flatMap((eachSet, idx) => {
    if (eachSet.length === 0) return '';
    if (eachSet.length > 1) {
      return `(${eachSet.join(` ${groups[idx].joinedBy} `)})`
    } else {
      return eachSet[0]
    }
  })

  return `?$filter=${filterStrings.join(` ${joinedBy} `)}`
}

function processFilters(group: IFilterGroup): string[] {
  if (group.items.length < 1) {
    return [];
  }

  return group.items.map(item => {
    if (item.type === 'group') {
      const joinOp = item.joinedBy || 'or';
      const nestedFilters = processFilters(item)
      if (nestedFilters.length > 1) {
        return `(${nestedFilters.join(` ${joinOp} `)})`
      } else {
        return nestedFilters[0]
      }
    } else {
      const val = item.filterBy.type === 'string' ? `'${item.filterValue}'` : item.filterValue
      return `${item.filterBy.value} ${item.filterOperator} ${val}`
    }
  })
}
