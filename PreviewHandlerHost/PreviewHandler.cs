using Serilog;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using Vanara.PInvoke;
using static Vanara.PInvoke.Shell32;
using static Vanara.PInvoke.ShlwApi;

namespace PreviewHandlerHost
{
    internal class PreviewHandler : IDisposable
    {
        private readonly IPreviewHandler _previewHandler;
        private readonly IPreviewHandlerVisuals _previewHandlerVisuals;
        //private readonly IOleWindow _oleWindow;

        public PreviewHandler(IPreviewHandlerFrame previewHandlerFrame, string fileName)
        {
            _previewHandler = CreatePreviewHandler(fileName);
            _previewHandlerVisuals = _previewHandler as IPreviewHandlerVisuals;
            
            var objectWithSite = _previewHandler as IObjectWithSite;
            Log.Debug($"objectWithSite: {objectWithSite}");
            objectWithSite?.SetSite(previewHandlerFrame).ThrowIfFailed("IObjectWithSite::SetSite");
        }

        private static IPreviewHandler CreatePreviewHandler(string fileName)
        {
            var clsid = GetPreviewHandlerCLSID(fileName);
            Log.Debug($"clsid: {clsid}");

            if (clsid != Guid.Empty)
            {
                var type = Type.GetTypeFromCLSID(clsid);
                Log.Debug($"type: {type}");

                if (type != null)
                {
                    // Adobe Reader bug?
                    var hr = Ole32.CoCreateInstance(clsid, null, Ole32.CLSCTX.CLSCTX_LOCAL_SERVER, typeof(IPreviewHandler).GUID, out var pph);
                    Log.Debug($"CoCreateInstance: {hr}");
                    //var pph = Activator.CreateInstance(type);
                    //Log.Debug($"CreateInstance: {pph}");

                    if (pph is IPreviewHandler previewHandler)
                    {
                        if (InitPreviewHandler(fileName, previewHandler))
                        {
                            return previewHandler;
                        }
                    }
                }
            }

            return null;
        }

        private static Guid GetPreviewHandlerCLSID(string fileName)
        {
            var flags = ASSOCF.ASSOCF_NOTRUNCATE | ASSOCF.ASSOCF_INIT_DEFAULTTOSTAR;
            // ASSOCSTR_SHELLEXTENSION
            // https://github.com/dahall/Vanara/pull/180
            var str = (ASSOCSTR)16; 
            var pszAssoc = Path.GetExtension(fileName);
            var pszExtra = typeof(IPreviewHandler).GUID.ToString("B");
            StringBuilder pszOut = null;
            uint pcchOut = 0;

            AssocQueryString(flags, str, pszAssoc, pszExtra, pszOut, ref pcchOut).ThrowIfFailed("GetPreviewHandlerCLSID::AssocQueryString");
            if (pcchOut == 0) return Guid.Empty;

            pszOut = new StringBuilder((int)pcchOut);
            AssocQueryString(flags, str, pszAssoc, pszExtra, pszOut, ref pcchOut).ThrowIfFailed("GetPreviewHandlerCLSID::AssocQueryString");

            return new Guid(pszOut.ToString());
        }

        private static bool InitPreviewHandler(string fileName, IPreviewHandler previewHandler)
        {
            return InitializeWithStream(fileName, previewHandler) ||
                   InitializeWithFile(fileName, previewHandler) ||
                   InitializeWithItem(fileName, previewHandler);
        }

        private static bool InitializeWithStream(string pszFile, IPreviewHandler previewHandler)
        {
            if (previewHandler is IInitializeWithStream initializeWithStream)
            {
                SHCreateStreamOnFile(pszFile, STGM.STGM_READ, out var ppstm).ThrowIfFailed("InitializeWithStream::SHCreateStreamOnFile");
                initializeWithStream.Initialize(ppstm, STGM.STGM_READ).ThrowIfFailed("InitializeWithStream::Initialize");
                Log.Debug("PreviewHandler::IInitializeWithStream");
                return true;
            }

            return false;
        }

        private static bool InitializeWithFile(string pszFile, IPreviewHandler previewHandler)
        {
            if (previewHandler is IInitializeWithFile initializeWithFile)
            {
                initializeWithFile.Initialize(pszFile, STGM.STGM_READ).ThrowIfFailed("InitializeWithFile::Initialize");
                Log.Debug("PreviewHandler::InitializeWithFile");
                return true;
            }

            return false;
        }

        private static bool InitializeWithItem(string pszFile, IPreviewHandler previewHandler)
        {
            if (previewHandler is IInitializeWithItem initializeWithItem)
            {
                var psi = ShellUtil.GetShellItemForPath(pszFile);
                initializeWithItem.Initialize(psi, STGM.STGM_READ).ThrowIfFailed("InitializeWithItem::Initialize");
                Log.Debug("PreviewHandler::InitializeWithItem");
                return true;
            }

            return false;
        }

        public bool SetFocus()
        {
            return _previewHandler.SetFocus().Succeeded;
        }

        public bool SetRect(Rect rect)
        {
            var prc = new RECT((int)rect.Left, (int)rect.Top, (int)rect.Right, (int)rect.Bottom);
            return _previewHandler.SetRect(prc).Succeeded;
        }

        public bool SetWindow(Window window, Rect rect)
        {
            var hwnd = new WindowInteropHelper(window).Handle;
            var prc = new RECT((int)rect.Left, (int)rect.Top, (int)rect.Right, (int)rect.Bottom);
            return _previewHandler.SetWindow(hwnd, prc).Succeeded;

        }

        public bool DoPreview()
        {
            var hr = _previewHandler.DoPreview();
            Log.Debug($"PreviewHandler::DoPreview: {hr}");
            return hr.Succeeded;
        }

        public bool SetBackgroundColor(Color color)
        {
            var hr = _previewHandlerVisuals?.SetBackgroundColor(color);
            return hr.HasValue && hr.Value.Succeeded;
        }

        public bool SetFont(Font font)
        {
            var logFont = new LOGFONT();
            font.ToLogFont(logFont);

            var hr = _previewHandlerVisuals?.SetFont(logFont);
            return hr.HasValue && hr.Value.Succeeded;
        }

        public bool SetTextColor(Color color)
        {
            var hr = _previewHandlerVisuals?.SetTextColor(color);
            return hr.HasValue && hr.Value.Succeeded;
        }

        public void Dispose()
        {
            if (_previewHandler != null)
            {
                try
                {
                    _previewHandler.Unload();
                }
                finally
                {
                    Marshal.FinalReleaseComObject(_previewHandler);
                }
            }
        }
    }
}
