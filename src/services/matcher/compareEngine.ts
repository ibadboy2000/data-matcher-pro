/**
 * 比对配置项接口定义
 */
export interface MatchOptions {
  keyIndicesA: number[];      // 主表联合主键的列索引
  keyIndicesB: number[];      // 次表联合主键的列索引
  compareColA?: number;       // 主表中需要核对的列索引
  compareColB?: number;       // 次表中需要核对的列索引
  backfillColB?: number;      // 次表中需要回填到主表的列索引
}

/**
 * 比对结果接口定义
 */
export interface MatchResult {
  onlyInARowIndices: number[]; // 仅主表存在的行索引
  onlyInBRowIndices: number[]; // 仅次表存在的行索引
  mismatchCells: Array<{       // 主键匹配但核对列不一致的单元格信息
    rowA: number;
    colA: number;
    valueA: any;
    valueB: any;
  }>;
  backfillData: Map<number, any>; // 需要回填的数据
}

export class DataMatcherEngine {
  private static KEY_SEPARATOR = '§_§';

  private static generateHashKey(rowData: any[], keyIndices: number[]): string {
    return keyIndices
      .map(index => {
        const val = rowData[index];
        return val !== null && val !== undefined ? String(val).trim() : '';
      })
      .join(this.KEY_SEPARATOR);
  }

  public static compare(
    tableA: any[][],
    tableB: any[][],
    options: MatchOptions,
    startRow: number = 1
  ): MatchResult {
    const result: MatchResult = {
      onlyInARowIndices: [],
      onlyInBRowIndices: [],
      mismatchCells: [],
      backfillData: new Map<number, any>()
    };

    const tableBMap = new Map<string, { rowIndex: number; rowData: any[] }>();
    
    for (let j = startRow; j < tableB.length; j++) {
      const rowB = tableB[j];
      if (!rowB || rowB.length === 0) continue;
      
      const keyB = this.generateHashKey(rowB, options.keyIndicesB);
      if (!tableBMap.has(keyB)) {
        tableBMap.set(keyB, { rowIndex: j, rowData: rowB });
      }
    }

    const matchedBRowIndices = new Set<number>();

    for (let i = startRow; i < tableA.length; i++) {
      const rowA = tableA[i];
      if (!rowA || rowA.length === 0) continue;

      const keyA = this.generateHashKey(rowA, options.keyIndicesA);

      if (tableBMap.has(keyA)) {
        const matchB = tableBMap.get(keyA)!;
        matchedBRowIndices.add(matchB.rowIndex);

        if (options.compareColA !== undefined && options.compareColB !== undefined) {
          const valA = rowA[options.compareColA];
          const valB = matchB.rowData[options.compareColB];

          const strA = valA !== null && valA !== undefined ? String(valA).trim() : '';
          const strB = valB !== null && valB !== undefined ? String(valB).trim() : '';

          if (strA !== strB) {
            result.mismatchCells.push({
              rowA: i,
              colA: options.compareColA,
              valueA: valA,
              valueB: valB
            });
          }
        }

        if (options.backfillColB !== undefined) {
          const backfillValue = matchB.rowData[options.backfillColB];
          result.backfillData.set(i, backfillValue);
        }
      } else {
        result.onlyInARowIndices.push(i);
      }
    }

    for (let j = startRow; j < tableB.length; j++) {
      if (!matchedBRowIndices.has(j)) {
        result.onlyInBRowIndices.push(j);
      }
    }

    return result;
  }
}
