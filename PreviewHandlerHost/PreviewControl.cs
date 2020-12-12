using Serilog;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Vanara.PInvoke;
using static Vanara.PInvoke.Shell32;

namespace PreviewHandlerHost
{
    [SuppressMessage("Design", "CA1001:Типы, владеющие высвобождаемыми полями, должны быть высвобождаемыми", Justification = "<Ожидание>")]
    public sealed class PreviewControl : Control, IPreviewHandlerFrame
    {
        private PreviewHandler _previewHandler;
        private string _fileName;

        public PreviewControl()
        {
            File.Delete("log.txt");
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Async(a => a.File("log.txt"))
                .CreateLogger();

            Loaded += PreviewControl_Loaded;
        }

        private void PreviewControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (_fileName != null)
            {
                ShowPreview();
            }
        }

        protected override void OnVisualParentChanged(DependencyObject oldParent)
        {
            if (Parent == null)
            {
                _previewHandler?.Dispose();
                _previewHandler = null;
            }

            base.OnVisualParentChanged(oldParent);
        }

        HRESULT IPreviewHandlerFrame.GetWindowContext(out PREVIEWHANDLERFRAMEINFO pinfo)
        {
            pinfo = new PREVIEWHANDLERFRAMEINFO();
            Log.Debug("GetWindowContext");
            return HRESULT.E_NOTIMPL;
        }

        HRESULT IPreviewHandlerFrame.TranslateAccelerator(in MSG pmsg)
        {
            Log.Debug("TranslateAccelerator");
            return HRESULT.E_NOTIMPL;
        }

        public string FileName
        {
            get { return (string)GetValue(FileNameProperty); }
            set { SetValue(FileNameProperty, value); }
        }

        public static readonly DependencyProperty FileNameProperty =
            DependencyProperty.Register(nameof(FileName), typeof(string), typeof(PreviewControl), new PropertyMetadata(null, OnFileNameChanged));

        private static void OnFileNameChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var previewControl = (PreviewControl)d;
            previewControl.SetFileName((string)e.NewValue);
        }

        private void SetFileName(string fileName)
        {
            _previewHandler?.Dispose();
            _previewHandler = null;
            _fileName = null;

            if (!string.IsNullOrWhiteSpace(fileName))
            {
                _fileName = fileName;
                if (IsLoaded)
                {
                    ShowPreview();
                }
            }
        }

        private void ShowPreview()
        {
            if (_previewHandler == null)
            {
                _previewHandler = new PreviewHandler(this, _fileName);
                if (Window.GetWindow(this) is Window window)
                {
                    var rect = TransformToVisual((Visual)Parent).TransformBounds(new Rect(RenderSize));
                    Log.Debug($"Rect: {rect}");
                    Log.Debug($"SetWindow: {_previewHandler.SetWindow(window, rect)}");
                    Log.Debug($"SetRect: {_previewHandler.SetRect(rect)}");
                    Log.Debug($"DoPreview: {_previewHandler.DoPreview()}");
                }
            }
        }
    }
}
