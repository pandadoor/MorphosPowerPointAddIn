using System;
using System.Runtime.InteropServices;

namespace MorphosPowerPointAddIn.Utilities
{
    internal sealed class OfficeBusyMessageFilter : IDisposable
    {
        private const int ServerCallIsHandled = 0;
        private const int ServerCallRetryLater = 2;
        private const int PendingMessageWaitDefProcess = 2;
        private const int RetryDelayMilliseconds = 150;
        private const int MaximumRetryWindowMilliseconds = 30000;

        private readonly RetryMessageFilter _currentFilter;
        private readonly IntPtr _currentFilterPointer;
        private readonly IntPtr _previousFilterPointer;
        private bool _disposed;
        private static readonly IDisposable NoOpRegistration = new NoOpDisposable();

        private OfficeBusyMessageFilter(RetryMessageFilter currentFilter, IntPtr currentFilterPointer, IntPtr previousFilterPointer)
        {
            _currentFilter = currentFilter;
            _currentFilterPointer = currentFilterPointer;
            _previousFilterPointer = previousFilterPointer;
        }

        public static IDisposable Register()
        {
            RetryMessageFilter currentFilter = null;
            var currentFilterPointer = IntPtr.Zero;

            try
            {
                currentFilter = new RetryMessageFilter();
                currentFilterPointer = Marshal.GetComInterfaceForObject(currentFilter, typeof(IMessageFilter));
                IntPtr previousFilterPointer;
                CoRegisterMessageFilter(currentFilterPointer, out previousFilterPointer);
                return new OfficeBusyMessageFilter(currentFilter, currentFilterPointer, previousFilterPointer);
            }
            catch
            {
                ReleasePointer(currentFilterPointer);
                return NoOpRegistration;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            IntPtr currentFilterPointer;
            CoRegisterMessageFilter(_previousFilterPointer, out currentFilterPointer);
            ReleasePointer(currentFilterPointer);
            ReleasePointer(_previousFilterPointer);
            ReleasePointer(_currentFilterPointer);
        }

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IntPtr newFilter, out IntPtr oldFilter);

        private static void ReleasePointer(IntPtr pointer)
        {
            if (pointer == IntPtr.Zero)
            {
                return;
            }

            try
            {
                Marshal.Release(pointer);
            }
            catch
            {
            }
        }

        [ComImport]
        [ComVisible(true)]
        [Guid("00000016-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType);

            [PreserveSig]
            int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType);
        }

        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.None)]
        private sealed class RetryMessageFilter : IMessageFilter
        {
            public int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
            {
                return ServerCallIsHandled;
            }

            public int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType)
            {
                if (dwRejectType == ServerCallRetryLater && dwTickCount < MaximumRetryWindowMilliseconds)
                {
                    return RetryDelayMilliseconds;
                }

                return -1;
            }

            public int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType)
            {
                return PendingMessageWaitDefProcess;
            }
        }

        private sealed class NoOpDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
