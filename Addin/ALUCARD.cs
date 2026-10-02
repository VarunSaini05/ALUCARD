using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;

namespace ALUCARD
{
    [ComVisible(true)]
    [Guid("7E8D4B8A-3D52-4B74-9B7A-5F4E8C2D1A61")]
    [ProgId("ALUCARD.SwAddin")]
    public class SwAddin : ISwAddin
    {
        private ISldWorks _swApp;
        private int _cookie;

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            _swApp = (ISldWorks)ThisSW;
            _cookie = Cookie;

            return true;
        }

        public bool DisconnectFromSW()
        {
            _swApp = null;
            _cookie = 0;

            return true;
        }
    }
}
