using Shell32;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RecycleTray
{
    public sealed class ComObject<T> : IDisposable
    {
        public readonly Shell Shell;
        public readonly T Object;

        public ComObject(
            Shell shell,
            T comObject)
        {
            Shell = shell;
            Object = comObject;
        }

        public void Dispose()
        {
            Marshal.FinalReleaseComObject(Shell);
        }
    }

    public static class TrashBin
    {
        public static void OpenExplorer()
        {
            Process.Start("explorer.exe", "shell:RecycleBinFolder");
        }

        public static ComObject<Folder> GetTrashFolder()
        {
            var shell = new Shell();
            Folder trash = shell.NameSpace(0x000a);

            return new ComObject<Folder>(shell, trash);
        }

        private enum RecycleFlag : int
        {
            SHERB_NOCONFIRMATION = 0x00000001, // No confirmation, when emptying
        }

        [DllImport("Shell32.dll")]
        private static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, RecycleFlag dwFlags);

        public static void EmptyRecycleBin()
        {
            SHEmptyRecycleBin(IntPtr.Zero, null, RecycleFlag.SHERB_NOCONFIRMATION);
        }
    }
}
