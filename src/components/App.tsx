import * as React from 'react';
import { useState, useEffect } from 'react';
import { 
  Stack, 
  Text, 
  Dropdown, 
  IDropdownOption, 
  Checkbox, 
  PrimaryButton, 
  DefaultButton, 
  MessageBar, 
  MessageBarType, 
  Separator,
  Spinner,
  SpinnerSize,
  initializeIcons
} from '@fluentui/react';
import { ExcelMatchService } from '../services/excel/matchService';
import { MatchOptions } from '../services/matcher/compareEngine';

initializeIcons();

export const App: React.FC = () => {
  const [loading, setLoading] = useState<boolean>(false);
  const [sheets, setSheets] = useState<string[]>([]);
  const [columnsA, setColumnsA] = useState<string[]>([]);
  const [columnsB, setColumnsB] = useState<string[]>([]);

  const [selectedSheetA, setSelectedSheetA] = useState<string>('');
  const [selectedSheetB, setSelectedSheetB] = useState<string>('');
  const [selectedKeysA, setSelectedKeysA] = useState<number[]>([]);
  const [selectedKeysB, setSelectedKeysB] = useState<number[]>([]);
  const [compareColA, setCompareColA] = useState<number | undefined>(undefined);
  const [compareColB, setCompareColB] = useState<number | undefined>(undefined);
  const [backfillColB, setBackfillColB] = useState<number | undefined>(undefined);

  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);

  useEffect(() => {
    loadSheets();
  }, []);

  useEffect(() => {
    if (selectedSheetA) loadColumns(selectedSheetA, 'A');
  }, [selectedSheetA]);

  useEffect(() => {
    if (selectedSheetB) loadColumns(selectedSheetB, 'B');
  }, [selectedSheetB]);

  const loadSheets = async () => {
    setLoading(true);
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();
        const names = sheets.items.map(s => s.name);
        setSheets(names);
      });
    } catch (error) {
      showMessage('获取工作表失败', MessageBarType.error);
    } finally {
      setLoading(false);
    }
  };

  const loadColumns = async (sheetName: string, type: 'A' | 'B') => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const range = sheet.getUsedRange().getRow(0);
        range.load('values');
        await context.sync();
        
        const cols = range.values[0].map(v => String(v));
        if (type === 'A') setColumnsA(cols);
        else setColumnsB(cols);
      });
    } catch (error) {
      showMessage(`获取${type}表列名失败`, MessageBarType.error);
    }
  };

  const showMessage = (text: string, type: MessageBarType = MessageBarType.info) => {
    setMessage({ text, type });
  };

  const handleMatch = async () => {
    if (!selectedSheetA || !selectedSheetB) {
      showMessage('请先选择主表和次表！', MessageBarType.warning);
      return;
    }
    if (selectedKeysA.length === 0 || selectedKeysB.length === 0) {
      showMessage('请为两张表至少勾选一个联合主键！', MessageBarType.warning);
      return;
    }
    if (selectedKeysA.length !== selectedKeysB.length) {
      showMessage('主表和次表的主键列数量必须一致！', MessageBarType.warning);
      return;
    }

    setLoading(true);
    setMessage(null);

    const options: MatchOptions = {
      keyIndicesA: selectedKeysA,
      keyIndicesB: selectedKeysB,
      compareColA,
      compareColB,
      backfillColB
    };

    try {
      const result = await ExcelMatchService.executeMatch(selectedSheetA, selectedSheetB, options);
      if (result.success) {
        showMessage(result.message, MessageBarType.success);
      } else {
        showMessage(result.message, MessageBarType.error);
      }
    } catch (error: any) {
      showMessage(`执行失败: ${error.message}`, MessageBarType.error);
    } finally {
      setLoading(false);
    }
  };

  const sheetOptions: IDropdownOption[] = sheets.map(s => ({ key: s, text: s }));
  const getColOptions = (cols: string[]): IDropdownOption[] => 
    cols.map((c, i) => ({ key: i, text: `${i+1}. ${c}` }));

  return (
    <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20, height: '100%', backgroundColor: '#F3F2F1' }}>
      <Stack>
        <Text variant="xxLarge" style={{ color: '#106EBE', fontWeight: 'bold' }}>DataMatcher Pro</Text>
        <Text variant="small" style={{ color: '#605E5C' }}>海量数据极速比对专家</Text>
      </Stack>

      <Separator />

      <Stack tokens={{ childrenGap: 10 }} style={{ backgroundColor: 'white', padding: 15, borderRadius: 4 }}>
        <Text variant="mediumPlus" style={{ fontWeight: 'bold' }}>1. 选择工作表</Text>
        <Dropdown
          label="主表 (Sheet A)"
          selectedKey={selectedSheetA}
          options={sheetOptions}
          onChange={(_, opt) => setSelectedSheetA(opt?.key as string)}
          placeholder="请选择主表"
        />
        <Dropdown
          label="次表 (Sheet B)"
          selectedKey={selectedSheetB}
          options={sheetOptions}
          onChange={(_, opt) => setSelectedSheetB(opt?.key as string)}
          placeholder="请选择次表"
        />
      </Stack>

      {selectedSheetA && selectedSheetB && (
        <Stack tokens={{ childrenGap: 10 }} style={{ backgroundColor: 'white', padding: 15, borderRadius: 4 }}>
          <Text variant="mediumPlus" style={{ fontWeight: 'bold' }}>2. 勾选联合主键</Text>
          <Text variant="small" color="#605E5C">请确保两表勾选的列数一致，且按顺序对应</Text>
          
          <Stack horizontal tokens={{ childrenGap: 20 }}>
            <Stack tokens={{ childrenGap: 5 }}>
              <Text style={{ fontWeight: 'bold' }}>主表主键：</Text>
              {columnsA.map((col, index) => (
                <Checkbox 
                  key={`keyA_${index}`} 
                  label={col} 
                  onChange={(_, checked) => {
                    setSelectedKeysA(prev => checked ? [...prev, index] : prev.filter(i => i !== index));
                  }}
                />
              ))}
            </Stack>
            
            <Stack tokens={{ childrenGap: 5 }}>
              <Text style={{ fontWeight: 'bold' }}>次表主键：</Text>
              {columnsB.map((col, index) => (
                <Checkbox 
                  key={`keyB_${index}`} 
                  label={col} 
                  onChange={(_, checked) => {
                    setSelectedKeysB(prev => checked ? [...prev, index] : prev.filter(i => i !== index));
                  }}
                />
              ))}
            </Stack>
          </Stack>
        </Stack>
      )}

      {(columnsA.length > 0 || columnsB.length > 0) && (
        <Stack tokens={{ childrenGap: 10 }} style={{ backgroundColor: 'white', padding: 15, borderRadius: 4 }}>
          <Text variant="mediumPlus" style={{ fontWeight: 'bold' }}>3. 高级核对与回填 (选填)</Text>
          
          <Dropdown
            label="主表核对列"
            options={getColOptions(columnsA)}
            onChange={(_, opt) => setCompareColA(opt?.key as number)}
            placeholder="不一致将标橙"
          />
          <Dropdown
            label="次表核对列"
            options={getColOptions(columnsB)}
            onChange={(_, opt) => setCompareColB(opt?.key as number)}
            placeholder="与主表核对列比对"
          />
          <Dropdown
            label="次表回填列"
            options={getColOptions(columnsB)}
            onChange={(_, opt) => setBackfillColB(opt?.key as number)}
            placeholder="匹配成功后追加到主表末尾"
          />
        </Stack>
      )}

      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton 
          text="开始比对" 
          onClick={handleMatch} 
          disabled={loading} 
          iconProps={{ iconName: 'SearchAndApps' }}
          style={{ flexGrow: 2 }}
        />
        <DefaultButton 
          text="重置" 
          onClick={() => window.location.reload()} 
          disabled={loading}
          style={{ flexGrow: 1 }}
        />
      </Stack>

      {loading && (
        <Spinner size={SpinnerSize.large} label="正在极速比对中，请稍候..." />
      )}

      {message && (
        <MessageBar 
          messageBarType={message.type} 
          onDismiss={() => setMessage(null)}
          dismissButtonAriaLabel="Close"
        >
          <span style={{ whiteSpace: 'pre-line' }}>{message.text}</span>
        </MessageBar>
      )}
    </Stack>
  );
};
