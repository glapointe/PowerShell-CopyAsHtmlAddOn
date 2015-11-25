using System;
using System.Windows.Threading;

namespace Falchion.PowerShell.CopyAsHtmlAddOn.Utilities
{
    public sealed class DispatchHelper
    {
        private readonly Dispatcher _dispatcher;

        /// <summary>
        /// Dispatch helper.
        /// </summary>
        /// <param name="method">The method.</param>
        /// <returns></returns>
        public static object DispatchInvoke(DispatcherOperationCallback method)
        {
            return DispatchInvoke(System.Windows.Application.Current.Dispatcher, method);
        }

        public static object DispatchInvoke(Dispatcher dispatcher, DispatcherOperationCallback method)
        {
            return dispatcher.Invoke(DispatcherPriority.Send, method, null);
        }

        public DispatchHelper(DispatcherObject obj)
        {
            _dispatcher = obj.Dispatcher;
        }

        public bool DispatchRequired
        {
            get { return !_dispatcher.CheckAccess(); }
        }

        public void Invoke(Action action)
        {
            Invoke(DispatcherPriority.Normal, action);
        }

        public void Invoke(DispatcherPriority priority, Action action)
        {
            if (DispatchRequired)
            {
                _dispatcher.Invoke(priority, action);
            }
            else
            {
                action();
            }
        }

        public T Invoke<T>(Func<T> func)
        {
            return Invoke<T>(DispatcherPriority.Normal, func);
        }

        public T Invoke<T>(DispatcherPriority priority, Func<T> func)
        {
            if (DispatchRequired)
            {
                return (T)_dispatcher.Invoke(priority, func);
            }

            return func();
        }
    }
}