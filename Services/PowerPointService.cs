using System.Runtime.InteropServices;

namespace com_api.Services
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class PowerPointService : IDisposable
    {
        private dynamic _powerPointApp;
        private dynamic _currentPresentation;
        private bool _disposed = false;
        private bool _isInitialized = false;

        public PowerPointService()
        {
            InitializePowerPoint();
        }

        private bool InitializePowerPoint()
        {
            try
            {
                if (_powerPointApp == null && !_isInitialized)
                {
                    Type? ppType = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (ppType == null)
                    {
                        Console.WriteLine("PowerPoint is not installed on this machine.");
                        return false;
                    }

                    try
                    {
                        _powerPointApp = Activator.CreateInstance(ppType);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"PowerPoint initialization failed: {ex.Message}");
                        return false;
                    }

                    if (_powerPointApp == null)
                    {
                        Console.WriteLine("Failed to create PowerPoint application instance.");
                        return false;
                    }

                    _powerPointApp.DisplayAlerts = false;
                    _isInitialized = true;
                    return true;
                }
                return _isInitialized;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing PowerPoint: {ex.Message}");
                return false;
            }
        }

        public bool OpenPresentation(string filePath, bool startSlideShow = true)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    Console.WriteLine($"File does not exist: {filePath}");
                    return false;
                }

                if (!_isInitialized && !InitializePowerPoint())
                {
                    return false;
                }

                if (_currentPresentation != null)
                {
                    try
                    {
                        _currentPresentation.Close();
                        _currentPresentation = null;
                    }
                    catch { }
                }

                var presentations = _powerPointApp?.Presentations;
                if (presentations == null)
                {
                    return false;
                }

                _currentPresentation = presentations.Open(filePath);

                if (startSlideShow && _currentPresentation != null)
                {
                    var settings = _currentPresentation.SlideShowSettings;
                    if (settings != null)
                    {
                        settings.ShowType = 1; // ppShowTypeSpeaker
                        settings.Run();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening presentation: {ex.Message}");
                return false;
            }
        }

        public bool GoToSlide(int slideNumber)
        {
            try
            {
                if (_powerPointApp == null || _currentPresentation == null)
                {
                    return false;
                }

                dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                if (slideShowWindows == null || Convert.ToInt32(slideShowWindows.Count) <= 0)
                {
                    return false;
                }

                dynamic? slides = _currentPresentation.Slides;
                if (slides == null)
                {
                    return false;
                }

                int slidesCount = Convert.ToInt32(slides.Count);
                if (slideNumber < 1 || slideNumber > slidesCount)
                {
                    return false;
                }

                dynamic? slideShow = slideShowWindows[1];
                if (slideShow == null)
                {
                    return false;
                }

                dynamic? view = slideShow.View;
                if (view == null)
                {
                    return false;
                }

                view.GotoSlide(slideNumber);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error navigating to slide {slideNumber}: {ex.Message}");
                return false;
            }
        }

        public bool ClosePresentation()
        {
            try
            {
                if (_currentPresentation == null)
                {
                    return false;
                }

                _powerPointApp.DisplayAlerts = 0;
                _currentPresentation.Saved = true;

                if (_powerPointApp != null)
                {
                    try
                    {
                        dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                        if (slideShowWindows != null && Convert.ToInt32(slideShowWindows.Count) > 0)
                        {
                            dynamic? slideShow = slideShowWindows[1];
                            if (slideShow?.View != null)
                            {
                                slideShow.View.Exit();
                            }
                        }
                    }
                    catch { }
                }

                _currentPresentation.Close();
                _currentPresentation = null;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error closing presentation: {ex.Message}");
                return false;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (_powerPointApp != null)
                {
                    try
                    {
                        _powerPointApp.DisplayAlerts = 0;
                    }
                    catch { }
                }

                if (_currentPresentation != null)
                {
                    try
                    {
                        _currentPresentation.Saved = true;
                        ClosePresentation();
                        Marshal.ReleaseComObject(_currentPresentation);
                    }
                    catch { }
                }

                if (_powerPointApp != null)
                {
                    try
                    {
                        _powerPointApp.Quit();
                        Marshal.ReleaseComObject(_powerPointApp);
                    }
                    catch { }
                }

                _currentPresentation = null;
                _powerPointApp = null;
                _disposed = true;
            }
        }

        ~PowerPointService()
        {
            Dispose(false);
        }
    }
}
