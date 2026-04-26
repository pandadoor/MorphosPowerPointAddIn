using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class DialogWindowHelper
    {
        public static void AttachToPowerPoint(Window dialog, PowerPoint.Application application)
        {
            if (dialog == null || application == null)
            {
                return;
            }

            try
            {
                new WindowInteropHelper(dialog)
                {
                    Owner = TryGetOwnerHandle(application)
                };
            }
            catch
            {
            }
        }

        public static System.Windows.Forms.IWin32Window TryGetPowerPointOwnerWindow()
        {
            var ownerHandle = TryGetOwnerHandle(Globals.ThisAddIn == null ? null : Globals.ThisAddIn.Application);
            return ownerHandle == IntPtr.Zero ? null : new PowerPointWindowHandle(ownerHandle);
        }

        public static System.Windows.Forms.IWin32Window TryGetOwnerWindow(Window dialog)
        {
            if (dialog != null)
            {
                try
                {
                    var dialogHandle = new WindowInteropHelper(dialog).Handle;
                    if (dialogHandle != IntPtr.Zero)
                    {
                        return new PowerPointWindowHandle(dialogHandle);
                    }
                }
                catch
                {
                }
            }

            return TryGetPowerPointOwnerWindow();
        }

        private static IntPtr TryGetOwnerHandle(PowerPoint.Application application)
        {
            if (application == null)
            {
                return IntPtr.Zero;
            }

            try
            {
                return new IntPtr(application.HWND);
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        private sealed class PowerPointWindowHandle : System.Windows.Forms.IWin32Window
        {
            public PowerPointWindowHandle(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }
    }
}
