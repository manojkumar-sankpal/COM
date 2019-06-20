using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SHDocVw;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace IEExtension
{
    [
        ComVisible(true),
        Guid("A13A7318-E265-4DEC-B798-B20EBB12042C"),
        ClassInterface(ClassInterfaceType.None)
    ]

    public class BHO : IObjectWithSite
    {
        private object m_pUnkSite;
        private IWebBrowser2 m_webBrowser2;
        private DWebBrowserEvents2_Event m_webBrowser2Events;

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            //
            //  TODO : Perform task after document is loaded completly 
            //
            //HTMLDocument messageBoxText = m_webBrowser2.Document;
            //System.Windows.Forms.MessageBox.Show("Hello World.....");
        }

        public void OnBeforeNavigate(object pDisp, ref object URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            string strAccessurl = (string)URL;

            if (strAccessurl.Contains("facebook.com"))
            {
                System.Windows.Forms.MessageBox.Show("Sorry !! You are not authorize to access Facebook.");
                Cancel = true;
            }
        }

        #region IObjectWithSite Interface Implementation

        public int SetSite([In, MarshalAs(UnmanagedType.IUnknown)]object pUnkSite)
        {
            if (pUnkSite != null)
            {
                m_pUnkSite = pUnkSite;
                m_webBrowser2 = (IWebBrowser2)pUnkSite;
                m_webBrowser2Events = (DWebBrowserEvents2_Event)pUnkSite;
                m_webBrowser2Events.DocumentComplete += OnDocumentComplete;
                m_webBrowser2Events.BeforeNavigate2 += OnBeforeNavigate;
            }
            else
            {
                m_webBrowser2Events.DocumentComplete -= OnDocumentComplete;
                m_webBrowser2Events.BeforeNavigate2 -= OnBeforeNavigate;
                m_pUnkSite = null;
            }

            return 0;
        }

        public int GetSite(ref Guid riid, out IntPtr ppvSite)
        {
            var pUnk = Marshal.GetIUnknownForObject(m_pUnkSite);

            try
            {
                return Marshal.QueryInterface(pUnk, ref riid, out ppvSite);
            }
            finally
            {
                Marshal.Release(pUnk);
            }
        }

        #endregion

        #region BHO Registration/Unregistration function

        //
        //  Administrator command prompt :
        //  a. Registration of COM Assembly
        //  cmd> regasm /codebase <DLL Path>
        //
        //  Run internt explorer, if shows BHO status as incompatible then, Goto Internet option-> advanced->security->"Enable enhanced protection mode" 
        //  Uncheck it. (This is happen due to DLL is not signed) 
        //
        //  b. Un-Registration
        //  cmd> regasm /u <DLL path>
        //

        public static string BHO_KEY = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHO_KEY, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(BHO_KEY);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
                ourKey = registryKey.CreateSubKey(guid);

            ourKey.SetValue("NoExplorer", 1);
            registryKey.Close();
            ourKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHO_KEY, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }

        #endregion

    } // End of class
}
