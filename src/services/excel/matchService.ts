import { DataMatcherEngine, MatchOptions, MatchResult } from '../matcher/compareEngine';

export class ExcelMatchService {
  private static COLORS = {
    onlyInA: '#E2EFDA', // 浅绿
    onlyInB: '#DDEBF7', // 浅蓝
    mismatch: '#FCE4D6' // 浅橙/红
  };

  public static async executeMatch(
    sheetNameA: string,
    sheetNameB: string,
    options: MatchOptions
  ): Promise<{ success: boolean; message: string }> {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheetA = sheets.getItem(sheetNameA);
      const sheetB = sheets.getItem(sheetNameB);

      const rangeA = sheetA.getUsedRange();
      const rangeB = sheetB.getUsedRange();

      rangeA.load(['values', 'rowCount', 'columnCount']);
      rangeB.load(['values', 'rowCount', 'columnCount']);

      await context.sync();

      const dataA = rangeA.values;
      const dataB = rangeB.values;

      if (!dataA || dataA.length <= 1) {
        return { success: false, message: '主表(Sheet A)没有足够的数据。' };
      }

      const startRow = 1; 
      const matchResult: MatchResult = DataMatcherEngine.compare(dataA, dataB, options, startRow);

      // 3.1 仅主表存在
      for (const rowIndex of matchResult.onlyInARowIndices) {
        const rowRange = rangeA.getRow(rowIndex);
        rowRange.format.fill.color = this.COLORS.onlyInA;
      }

      // 3.2 仅次表存在
      for (const rowIndex of matchResult.onlyInBRowIndices) {
        const rowRange = rangeB.getRow(rowIndex);
        rowRange.format.fill.color = this.COLORS.onlyInB;
      }

      // 3.3 核对列不一致
      for (const mismatch of matchResult.mismatchCells) {
        const cellRange = rangeA.getCell(mismatch.rowA, mismatch.colA);
        cellRange.format.fill.color = this.COLORS.mismatch;
      }

      // 4. 处理数据回填
      if (options.backfillColB !== undefined) {
        const nextColIndex = rangeA.columnCount;
        const backfillColumnValues: any[][] = [];
        
        backfillColumnValues.push([`回填数据_${dataB[0][options.backfillColB]}`]);

        for (let i = startRow; i < rangeA.rowCount; i++) {
          if (matchResult.backfillData.has(i)) {
            backfillColumnValues.push([matchResult.backfillData.get(i)]);
          } else {
            backfillColumnValues.push(['']);
          }
        }

        const newColumnRange = rangeA.getOffsetRange(0, nextColIndex).getAbsoluteResizedRange(rangeA.rowCount, 1);
        newColumnRange.values = backfillColumnValues;
      }

      await context.sync();

      const summaryMessage = `核对完成！
- 仅主表存在：${matchResult.onlyInARowIndices.length} 行 (已标绿)
- 仅次表存在：${matchResult.onlyInBRowIndices.length} 行 (已标蓝)
- 数据不一致：${matchResult.mismatchCells.length} 处 (已标橙)
- 数据回填：${options.backfillColB !== undefined ? '已追加至末尾新列' : '未选择'}`;

      return { success: true, message: summaryMessage };
    });
  }
}
