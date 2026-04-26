using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class ErrorReporter
    {
        public static void Show(string message, Exception exception = null)
        {
            var details = exception == null
                ? message
                : message + Environment.NewLine + Environment.NewLine + exception.Message;

            if (exception != null)
            {
                Trace.WriteLine(exception);
            }

            var owner = DialogWindowHelper.TryGetPowerPointOwnerWindow();
            if (owner != null)
            {
                MessageBox.Show(owner, details, "Morphos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show(details, "Morphos", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
