import * as React from 'react';
import * as ReactDOM from 'react-dom/client';
import { App } from './components/App';
import { mergeStyles } from '@fluentui/react';

// 注入全局样式，确保 body 没有默认边距
mergeStyles({
  ':global(body)': {
    margin: 0,
    padding: 0,
    overflow: 'hidden',
    height: '100vh'
  }
});

// 确保 Office.js 初始化完成后再挂载 React
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const rootElement = document.getElementById('container');
    if (rootElement) {
      const root = ReactDOM.createRoot(rootElement);
      root.render(
        <React.StrictMode>
          <App />
        </React.StrictMode>
      );
    }
  }
});
