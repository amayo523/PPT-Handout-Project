using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptNotesHandoutMaker.Core
{
    internal static class PowerPointInteropUtil
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(
            ref Guid rclsid,
            IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public static bool TryGetRunningPowerPoint(out PowerPoint.Application? app)
        {
            app = null;

            try
            {
                Guid clsidPowerPoint = new("91493441-5A91-11CF-8700-00AA0060263B");
                GetActiveObject(ref clsidPowerPoint, IntPtr.Zero, out object obj);
                app = (PowerPoint.Application)obj;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void FinalRelease(object? comObj)
        {
            if (comObj == null)
                return;

            if (!OperatingSystem.IsWindows())
                return;

            try { Marshal.FinalReleaseComObject(comObj); } catch { }
        }
    }
}