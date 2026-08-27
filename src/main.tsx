import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

class AppErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { hasError: boolean; message: string }
> {
  state = { hasError: false, message: '' };

  static getDerivedStateFromError(error: unknown) {
    return {
      hasError: true,
      message: error instanceof Error ? error.message : '页面发生了未知错误',
    };
  }

  componentDidCatch(error: unknown, info: React.ErrorInfo) {
    console.error('Tavern Vault UI error:', error, info);
  }

  handleReload = () => window.location.reload();

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center p-6 bg-zinc-50 text-zinc-900">
          <div className="w-full max-w-md rounded-2xl border border-zinc-200 bg-white p-6 shadow-sm">
            <h1 className="text-lg font-bold">页面遇到错误，但数据没有因此删除</h1>
            <p className="mt-2 text-sm text-zinc-500 break-words">{this.state.message}</p>
            <button
              type="button"
              onClick={this.handleReload}
              className="mt-5 rounded-xl bg-zinc-900 px-4 py-2 text-sm font-semibold text-white"
            >
              重新打开工具箱
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  </React.StrictMode>
);
