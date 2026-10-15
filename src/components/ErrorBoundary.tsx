import React, { Component, ErrorInfo, ReactNode } from 'react';
import { AlertTriangle, RefreshCw, Trash2 } from 'lucide-react';

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
  errorInfo: ErrorInfo | null;
}

export class ErrorBoundary extends Component<Props, State> {
  public state: State = {
    hasError: false,
    error: null,
    errorInfo: null,
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error, errorInfo: null };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error('Uncaught error in React tree:', error, errorInfo);
    this.setState({ errorInfo });
  }

  private handleResetAll = () => {
    if (window.confirm('آیا از بازنشانی داده‌های موقت اطمینان دارید؟ تمام داده‌های اولیه مجدداً بارگذاری می‌شوند.')) {
      localStorage.clear();
      window.location.reload();
    }
  };

  private handleReload = () => {
    window.location.reload();
  };

  public render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col items-center justify-center p-6 text-right" dir="rtl">
          <div className="bg-slate-800 border border-red-500/30 rounded-2xl max-w-xl w-full p-8 shadow-2xl">
            <div className="flex items-center gap-3 text-red-400 mb-4">
              <AlertTriangle className="w-8 h-8 flex-shrink-0" />
              <h1 className="text-xl font-bold">خطا در بارگذاری برنامه</h1>
            </div>
            
            <p className="text-slate-300 text-sm mb-4 leading-relaxed">
              متأسفانه مشکلی در اجرای برنامه رخ داده است. می‌توانید صفحه را مجدداً بارگذاری کنید یا داده‌های ذخیره‌شده را بازنشانی نمایید.
            </p>

            <div className="bg-slate-950/70 p-4 rounded-xl text-xs font-mono text-red-300 overflow-x-auto mb-6 border border-slate-700/50">
              {this.state.error?.toString()}
            </div>

            <div className="flex flex-wrap items-center gap-3 justify-end">
              <button
                onClick={this.handleResetAll}
                className="flex items-center gap-2 px-4 py-2.5 bg-red-600/20 text-red-400 hover:bg-red-600/30 rounded-xl transition text-sm font-medium border border-red-500/30"
              >
                <Trash2 className="w-4 h-4" />
                بازنشانی حافظه برنامه
              </button>
              <button
                onClick={this.handleReload}
                className="flex items-center gap-2 px-5 py-2.5 bg-blue-600 text-white hover:bg-blue-700 rounded-xl transition text-sm font-medium shadow-lg shadow-blue-500/20"
              >
                <RefreshCw className="w-4 h-4" />
                تلاش مجدد (بارگذاری دوباره)
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
