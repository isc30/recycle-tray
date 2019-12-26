using Microsoft.Win32;
using Shell32;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace RecycleTray
{
    static class Program
    {
        private static NotifyIcon _tray = null;
        private static ComObject<Folder> _trashFolder = null;

        [STAThread]
        public static void Main()
        {
            _trashFolder = TrashBin.GetTrashFolder();

            int isLightTheme = (int)Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\", "SystemUsesLightTheme", 0);

            string icon = isLightTheme == 1
                ? "light.png"
                : "dark.png";

            Bitmap bitmap = (Bitmap)Image.FromFile(icon);

            _tray = new NotifyIcon
            {
                Icon = Icon.FromHandle(bitmap.GetHicon()),
                Visible = true,
            };

            _tray.MouseClick += Icon_Click;

            var contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Open", null, MenuOpen);
            contextMenu.Items.Add("Empty Trash Bin", null, MenuEmpty);
            contextMenu.Items.Add("Exit", null, MenuExit);
            _tray.ContextMenuStrip = contextMenu;

            Timer timer = new Timer();
            timer.Interval = 5000;
            timer.Tick += (object sender, EventArgs ev) => UpdateTray();

            UpdateTray();
            timer.Start();

            Application.Run();
        }

        private static void UpdateTray()
        {
            int itemCount = _trashFolder.Object.Items().Count;

            if (itemCount > 0)
            {
                _tray.Visible = true;
                _tray.Text = $"{itemCount} items in the trash";
            }
            else
            {
                _tray.Visible = false;
            }
        }

        private static void MenuOpen(object sender, EventArgs ev)
        {
            TrashBin.OpenExplorer();
        }

        private static void MenuEmpty(object sender, EventArgs ev)
        {
            TrashBin.EmptyRecycleBin();
            UpdateTray();
        }

        private static void MenuExit(object sender, EventArgs ev)
        {
            _trashFolder.Dispose();
            _tray.Visible = false;
            Application.Exit();
        }

        private static void Icon_Click(object sender, MouseEventArgs ev)
        {
            switch (ev.Button)
            {
                case MouseButtons.Left:
                    TrashBin.OpenExplorer();
                    break;
            }
        }
    }
}
