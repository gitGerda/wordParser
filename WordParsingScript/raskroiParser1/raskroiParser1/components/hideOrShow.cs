using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace raskroiParser1.components
{
    public static class hideOrShow
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        public static int hideOrShowFunc(string[] args)
        {
            var handle = GetConsoleWindow();

            foreach (string s in args)
            {
                if (s == "--hide")
                {
                    ShowWindow(handle, SW_HIDE);
                    return 0;
                }
            }

            ShowWindow(handle, SW_SHOW);
            return 1;
        }
    }
}
