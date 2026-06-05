interface IFilter {
  filterBy: string,
  filterValue: string | number,
  filterOperator: 'eq' | 'ne' | 'gt' | 'ge' | 'lt' | 'le'
}
type JoinOption = 'and' | 'or' | 'not'

interface IFilterGroup {
  filters: IFilterGroup[] | IFilter[],
  joinedBy?: JoinOption
}

function buildFilterQuery(filters: IFilterGroup[], joinedBy: JoinOption = 'or') {
  if(filters.length < 1) {
    writeLogToSpreadsheet("Filter array is empty")
    return ""
  }
  const filterStrings = processFilters(filters)
   
  if (filterStrings.length === 0) return "";

  return `?$filter=${filterStrings.join(` ${joinedBy} `)}`
}

function isFilterArray(array: IFilterGroup[] | IFilter[]): array is IFilter[] {
  if(array.length === 0) return false;
  return 'filterBy' in array[0];
}

function processFilters(filters: IFilterGroup[]): string[] {
  if(filters.length < 1) {
    writeLogToSpreadsheet("Filter array is empty")
    return [];
  }
  
  return filters.flatMap(item => {
    if (item.filters.length === 0) return []; 
    const joinOp = item.joinedBy || 'or'; 

    if(isFilterArray(item.filters)) {
      const filterStrs = item.filters.map((filter) => {
        const val = typeof filter.filterValue === 'string' 
          ? `'${filter.filterValue}'` 
          : filter.filterValue;
          
        return `${filter.filterBy} ${filter.filterOperator} ${val}`;
      });
      if(filterStrs.length > 1) {
        return `(${filterStrs.join(` ${joinOp} `)})`
      } else {
        return filterStrs[0];
      }
    } else {
      const nestedStrs = processFilters(item.filters);
      if (nestedStrs.length > 1) {
        return `(${nestedStrs.join(` ${joinOp} `)})`;
      }
      return nestedStrs; 
    }
  })
}
