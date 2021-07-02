using System;
using System.Runtime.InteropServices;

namespace ExcelOrderAddIn
{
    public class ScrollingRichTextBox : System.Windows.Forms.RichTextBox
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(
            IntPtr hWnd,
            uint msg,
            IntPtr wParam,
            IntPtr lParam);

        private const int WmVScroll = 277;
        private const int SbBottom = 7;

        /**
         * Scrolls to the bottom of the RichTextBox.
         */
        public void ScrollToBottom()
        {
            SendMessage(Handle, WmVScroll, new IntPtr(SbBottom), new IntPtr(0));
        }
    }
}
