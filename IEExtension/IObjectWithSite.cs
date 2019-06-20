
using System;
using System.Runtime.InteropServices;

namespace IEExtension
{
    [
        ComImport,
        ComVisible(true),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")
    ]

    interface IObjectWithSite
    {
        [PreserveSig]
        int SetSite([In, MarshalAs(UnmanagedType.IUnknown)]object pUnkSite);

        [PreserveSig]
        int GetSite(ref Guid riid, out IntPtr ppvSite);
    }
}
