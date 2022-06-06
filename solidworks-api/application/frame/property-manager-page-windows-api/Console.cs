using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace InsertModelItemsWinAPI
{
    class Program
    {
        public class SearchData
        {
            public string ClassName;
            public string Title;
            public List<IntPtr> Results;

            public SearchData()
            {
                Results = new List<IntPtr>();
            }
        }

        #region Windows API

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);
        
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        public delegate bool WindowEnumProc(IntPtr hWnd, ref SearchData lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hWnd, WindowEnumProc func, ref SearchData lParam);

        #endregion

        static void Main(string[] args)
        {
            var app = (ISldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));

            int errs = -1;
            int warns = -1;

            var filePath = args.First();

            var doc = app.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errs, ref warns);
            
            const int WM_COMMAND = 0x0111;
            const int MODEL_ITEMS_CMD = 38374;

            //get the handle to SOLIDWORKS window
            var hWnd = new IntPtr(app.IFrameObject().GetHWnd());

            //open 'Model Items' property manager page
            SendMessage(hWnd, WM_COMMAND, MODEL_ITEMS_CMD, 0);

            var modelItemPageWnd = FindPropertyPageByName(hWnd, "Model Items");

            //Find the check box 'Include items from hidden features'
            var includeHiddenItemsWnd = FindWindows(modelItemPageWnd, "Include items from &hidden features", "Button").First();

            //check the found checkbox
            SetCheckBox(includeHiddenItemsWnd, true);

            //Find the source ComboBox (this is a first ComboBox in the page)
            var srcComboBox = FindWindows(modelItemPageWnd, "", "ComboBox").First();

            //Set the ComboBox selection to the first item (Entire Model)
            SetComboBox(srcComboBox, 0);

            const int swCommands_PmOK = -2;

            //Click OK on the PMPage to complete the operation
            app.RunCommand(swCommands_PmOK, "");

            doc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errs, ref warns);
            app.CloseDoc(doc.GetTitle());
        }

        private static void SetCheckBox(IntPtr checkBoxWnd, bool value)
        {
            const int BST_UNCHECKED = 0x0000;
            const int BST_CHECKED = 0x0001;
            const int BM_SETCHECK = 0x00F1;

            SendMessage(checkBoxWnd, BM_SETCHECK, value ? BST_CHECKED : BST_UNCHECKED, 0);
        }

        private static void SetComboBox(IntPtr comboBoxWnd, int index) 
        {
            const int CB_SETCURSEL = 0x014E;
            SendMessage(comboBoxWnd, CB_SETCURSEL, index, 0);
        }

        private static IntPtr FindPropertyPageByName(IntPtr swHwnd, string name)
        {
            var pagesWnd = FindWindows(swHwnd, "Dve sheet", "AfxWnd140u");

            foreach (var pageWnd in pagesWnd) 
            {
                if (FindWindows(pageWnd, name, "Button").Any()) 
                {
                    return pageWnd;
                }
            }

            throw new Exception($"Failed to find the property page '{name}'");
        }

        private static IntPtr[] FindWindows(IntPtr parentWnd, string title, string className)
        {
            var data = new SearchData()
            {
                ClassName = className,
                Title = title
            };

            var callbackProc = new WindowEnumProc(EnumChildWindowsCallback);
            EnumChildWindows(parentWnd, callbackProc, ref data);

            return data.Results.ToArray();
        }

        private static bool EnumChildWindowsCallback(IntPtr hWnd, ref SearchData data)
        {
            GetWindowInfo(hWnd, out string title, out string className);

            if ((string.IsNullOrEmpty(data.ClassName) || className == data.ClassName) && (string.IsNullOrEmpty(data.Title) || title == data.Title))
            {
                data.Results.Add(hWnd);
            }

            return true;
        }

        private static void GetWindowInfo(IntPtr hWnd, out string title, out string className)
        {
            var length = GetWindowTextLength(hWnd);
            var sb = new StringBuilder(length + 1);
            
            GetWindowText(hWnd, sb, sb.Capacity);

            title = sb.ToString();

            sb = new StringBuilder(256);
            
            GetClassName(hWnd, sb, sb.Capacity);

            className = sb.ToString().Trim();
        }
    }
}