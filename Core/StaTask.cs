using System;
using System.Threading;
using System.Threading.Tasks;

namespace PptNotesHandoutMaker
{
    public static class StaTask
    {
        public static Task Run(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));

            var tcs = new TaskCompletionSource<object?>();

            var thread = new Thread(() =>
            {
                try
                {
                    action();
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            thread.IsBackground = true;

            if (OperatingSystem.IsWindows())
                thread.SetApartmentState(ApartmentState.STA);

            thread.Start();


            return tcs.Task;
        }
        public static Task<T> Run<T>(Func<T> func)
        {
            if (func == null) throw new ArgumentNullException(nameof(func));

            var tcs = new TaskCompletionSource<T>();

            var thread = new Thread(() =>
            {
                try
                {
                    T result = func();
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            thread.IsBackground = true;

            if (OperatingSystem.IsWindows())
                thread.SetApartmentState(ApartmentState.STA);

            thread.Start();

            return tcs.Task;
        }
    }

}
