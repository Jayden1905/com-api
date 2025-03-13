using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace com_api.Services
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class PowerPointService : IDisposable
    {
        private dynamic? _powerPointApp;
        private dynamic? _currentPresentation;
        private bool _disposed = false;
        private bool _isInitialized = false;
        private Dictionary<string, string> _preloadedPresentations = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase
        );
        private string _lastOpenedPath = null;

        public PowerPointService()
        {
            // Pre-initialize PowerPoint immediately without the long delay
            // Just use a very short delay to let the service initialize basic components
            Task.Run(async () =>
            {
                // Reduced initial delay from 2000ms to 200ms
                await Task.Delay(200);

                // Aggressive initialization with minimal retries
                for (int attempt = 1; attempt <= 2; attempt++)
                {
                    try
                    {
                        Console.WriteLine($"Initializing PowerPoint (attempt {attempt})...");
                        if (InitializePowerPoint())
                        {
                            Console.WriteLine("PowerPoint initialized successfully.");
                            // Prepare the PowerPoint environment for faster subsequent operations
                            PrepareEnvironment();
                            break;
                        }

                        // Shorter backoff delay
                        await Task.Delay(500);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(
                            $"Error during PowerPoint initialization attempt {attempt}: {ex.Message}"
                        );
                        if (attempt < 2)
                            await Task.Delay(500);
                    }
                }
            });
        }

        // Add a new method to prepare the PowerPoint environment for better performance
        private void PrepareEnvironment()
        {
            try
            {
                if (_powerPointApp != null)
                {
                    // Set performance-related properties
                    try
                    {
                        // NOTE: Removed ScreenUpdating property as it's not available in all PowerPoint versions

                        // Set other performance options
                        dynamic options = _powerPointApp.Options;
                        if (options != null)
                        {
                            try
                            {
                                options.FeatureInstall = 0;
                            }
                            catch { } // msoFeatureInstallNone
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning during environment preparation: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error preparing PowerPoint environment: {ex.Message}");
            }
        }

        // Add diagnostic method to check PowerPoint and file access
        public string CheckPowerPointStatus(string filePath = null)
        {
            try
            {
                StringBuilder status = new StringBuilder();

                // Check if we have a PowerPoint instance
                status.AppendLine($"PowerPoint initialized: {_isInitialized}");
                status.AppendLine($"PowerPoint instance available: {_powerPointApp != null}");

                if (_powerPointApp != null)
                {
                    try
                    {
                        // Try to access PowerPoint properties
                        string version = _powerPointApp.Version?.ToString() ?? "unknown";
                        status.AppendLine($"PowerPoint version: {version}");
                        status.AppendLine($"PowerPoint visible: {_powerPointApp.Visible}");

                        // Check Presentations collection
                        var presentations = _powerPointApp.Presentations;
                        status.AppendLine(
                            $"Presentations collection: {(presentations != null ? "Available" : "Not available")}"
                        );

                        // Check current presentation
                        status.AppendLine(
                            $"Current presentation: {(_currentPresentation != null ? "Open" : "None")}"
                        );
                    }
                    catch (Exception ex)
                    {
                        status.AppendLine($"Error checking PowerPoint properties: {ex.Message}");
                    }
                }

                // Check file path if provided
                if (!string.IsNullOrEmpty(filePath))
                {
                    status.AppendLine($"\nChecking file: {filePath}");
                    status.AppendLine($"File exists: {File.Exists(filePath)}");

                    if (File.Exists(filePath))
                    {
                        try
                        {
                            // Get basic file info
                            var fileInfo = new FileInfo(filePath);
                            status.AppendLine($"File size: {fileInfo.Length} bytes");
                            status.AppendLine($"Last modified: {fileInfo.LastWriteTime}");
                            status.AppendLine($"Extension: {fileInfo.Extension}");

                            // Try to open the file for reading to check access permissions
                            using (var stream = File.OpenRead(filePath))
                            {
                                status.AppendLine("File readable: Yes");
                            }
                        }
                        catch (Exception ex)
                        {
                            status.AppendLine($"Error accessing file: {ex.Message}");
                        }
                    }
                }

                return status.ToString();
            }
            catch (Exception ex)
            {
                return $"Error in CheckPowerPointStatus: {ex.Message}";
            }
        }

        private bool InitializePowerPoint()
        {
            try
            {
                if (_powerPointApp == null && !_isInitialized)
                {
                    // Create PowerPoint application dynamically
                    Type? ppType = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (ppType == null)
                    {
                        Console.WriteLine("PowerPoint is not installed on this machine.");
                        return false;
                    }

                    _powerPointApp = Activator.CreateInstance(ppType);
                    if (_powerPointApp == null)
                    {
                        Console.WriteLine("Failed to create PowerPoint application instance.");
                        return false;
                    }

                    // NOTE: We keep PowerPoint visible from the start to avoid the error
                    // "Hiding the application window is not allowed"
                    // _powerPointApp.Visible = false; - This line caused the error

                    // Set performance optimizations for fast loading
                    try
                    {
                        // Disable alerts that might slow down operations
                        _powerPointApp.DisplayAlerts = false;

                        // Set startup options for faster performance
                        dynamic options = _powerPointApp.Options;
                        if (options != null)
                        {
                            // Set property values that might help performance
                            // These depend on PowerPoint version, so wrapped in try-catch
                            try
                            {
                                options.DisplayPasteOptions = false;
                            }
                            catch { }
                            try
                            {
                                options.FeatureInstall = 0;
                            }
                            catch { } // msoFeatureInstallNone
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log but continue - these are optional optimizations
                        Console.WriteLine($"Warning during PowerPoint optimization: {ex.Message}");
                    }

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

        public bool OpenPresentation(
            string filePath,
            bool startSlideShow = true,
            bool readOnly = true
        )
        {
            try
            {
                // Validate file path with minimal checks
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    Console.WriteLine($"File does not exist: {filePath}");
                    return false;
                }

                // Track this as last opened for potential quick reopen later
                _lastOpenedPath = Path.GetFullPath(filePath);

                // Check file extension - quick check
                string extension = Path.GetExtension(filePath).ToLower();
                if (extension != ".ppt" && extension != ".pptx" && extension != ".pptm")
                {
                    Console.WriteLine($"Unsupported format: {extension}");
                    return false;
                }

                // Ensure PowerPoint is initialized
                if (!_isInitialized && !InitializePowerPoint())
                {
                    return false;
                }

                // Pre-close any existing presentation
                if (_currentPresentation != null)
                {
                    try
                    {
                        _currentPresentation.Saved = true;
                        _currentPresentation.Close();
                        _currentPresentation = null;
                    }
                    catch (Exception)
                    {
                        // Ignore errors during close and continue
                        _currentPresentation = null;
                    }
                }

                // Get Presentations collection with retry
                var presentations = RetryOperation(() => _powerPointApp?.Presentations, null);
                if (presentations == null)
                {
                    Console.WriteLine("PowerPoint Presentations collection unavailable");
                    return false;
                }

                // Try to open the presentation with retry
                _currentPresentation = RetryOperation(
                    () =>
                    {
                        try
                        {
                            // Start with the simplest approach
                            return presentations.Open(filePath);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Simple open failed: {ex.Message}");

                            // Fall back to more parameters if needed
                            return presentations.Open(
                                filePath,
                                readOnly,
                                false, // Untitled
                                true // WithWindow
                            );
                        }
                    },
                    null
                );

                if (_currentPresentation == null)
                {
                    Console.WriteLine("Failed to open presentation after retry attempts");
                    return false;
                }

                // Start slideshow if requested
                if (startSlideShow)
                {
                    bool slideshowStarted = RetryOperation(
                        () =>
                        {
                            var settings = _currentPresentation.SlideShowSettings;
                            if (settings != null)
                            {
                                settings.ShowType = 1; // ppShowTypeSpeaker
                                settings.ShowWithAnimation = true;
                                settings.Run();
                                return true;
                            }
                            return false;
                        },
                        false
                    );

                    if (!slideshowStarted)
                    {
                        Console.WriteLine(
                            "Warning: Presentation opened but could not start slideshow"
                        );
                        // Continue anyway since the presentation is open
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

        public bool NextSlide()
        {
            try
            {
                // Validate PowerPoint app and check for active slideshow
                if (_powerPointApp == null)
                {
                    return false;
                }

                dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                if (slideShowWindows == null || Convert.ToInt32(slideShowWindows.Count) <= 0)
                {
                    return false;
                }

                // Get the first slideshow window (PowerPoint uses 1-based indexing)
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

                // Move to next slide - we've checked view is not null
                view.Next();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error navigating to next slide: {ex.Message}");
                return false;
            }
        }

        public bool PreviousSlide()
        {
            try
            {
                // Validate PowerPoint app and check for active slideshow
                if (_powerPointApp == null)
                {
                    return false;
                }

                dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                if (slideShowWindows == null || Convert.ToInt32(slideShowWindows.Count) <= 0)
                {
                    return false;
                }

                // Get the first slideshow window (PowerPoint uses 1-based indexing)
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

                // Move to previous slide - we've checked view is not null
                view.Previous();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error navigating to previous slide: {ex.Message}");
                return false;
            }
        }

        public bool GoToSlide(int slideNumber)
        {
            try
            {
                // Validate PowerPoint app and presentation
                if (_powerPointApp == null || _currentPresentation == null)
                {
                    return false;
                }

                dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                if (slideShowWindows == null || Convert.ToInt32(slideShowWindows.Count) <= 0)
                {
                    return false;
                }

                // Validate slide number
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

                // Get the first slideshow window (PowerPoint uses 1-based indexing)
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

                // Go to specific slide - we've checked view is not null
                view.GotoSlide(slideNumber);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error navigating to slide {slideNumber}: {ex.Message}");
                return false;
            }
        }

        public int GetCurrentSlideNumber()
        {
            try
            {
                // Validate PowerPoint app and check for active slideshow
                if (_powerPointApp == null)
                {
                    return -1;
                }

                dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                if (slideShowWindows == null || Convert.ToInt32(slideShowWindows.Count) <= 0)
                {
                    return -1;
                }

                // Get the first slideshow window
                dynamic? slideShow = slideShowWindows[1];
                if (slideShow == null)
                {
                    return -1;
                }

                dynamic? view = slideShow.View;
                if (view == null)
                {
                    return -1;
                }

                dynamic? slide = view.Slide;
                if (slide == null)
                {
                    return -1;
                }

                // Get slide number - we've checked slide is not null
                return Convert.ToInt32(slide.SlideNumber);
            }
            catch
            {
                return -1;
            }
        }

        public int GetTotalSlides()
        {
            try
            {
                if (_currentPresentation == null)
                {
                    return 0;
                }

                dynamic? slides = _currentPresentation.Slides;
                if (slides == null)
                {
                    return 0;
                }

                // Get slides count - we've checked slides is not null
                return Convert.ToInt32(slides.Count);
            }
            catch
            {
                return 0;
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

                // Prevent save prompts by explicitly marking the presentation as saved
                try
                {
                    // This explicitly tells PowerPoint not to prompt for saving
                    _powerPointApp.DisplayAlerts = 0; // ppAlertsNone = 0

                    // Mark the presentation as saved to prevent save dialog
                    _currentPresentation.Saved = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning when setting saved state: {ex.Message}");
                    // Continue anyway
                }

                // Exit slideshow if running
                if (_powerPointApp != null)
                {
                    try
                    {
                        dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                        if (slideShowWindows != null && Convert.ToInt32(slideShowWindows.Count) > 0)
                        {
                            dynamic? slideShow = slideShowWindows[1];
                            if (slideShow != null)
                            {
                                dynamic? view = slideShow.View;
                                if (view != null)
                                {
                                    // Exit the slideshow - we've checked view is not null
                                    view.Exit();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning when exiting slideshow: {ex.Message}");
                        // Continue anyway
                    }
                }

                // Close the presentation - we've checked _currentPresentation is not null
                try
                {
                    _currentPresentation.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during presentation close: {ex.Message}");
                    // Try a more aggressive approach if normal close fails
                    try
                    {
                        Marshal.ReleaseComObject(_currentPresentation);
                    }
                    catch
                    {
                        // Ignore errors during forced release
                    }
                }

                _currentPresentation = null;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error closing presentation: {ex.Message}");
                return false;
            }
        }

        // Add a public method to force-quit PowerPoint
        public bool ForceQuitPowerPoint()
        {
            try
            {
                Console.WriteLine("Force quitting PowerPoint...");

                // First close any open presentation
                if (_currentPresentation != null)
                {
                    try
                    {
                        // Prevent save dialogs
                        _currentPresentation.Saved = true;
                        _currentPresentation.Close();
                        Marshal.ReleaseComObject(_currentPresentation);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(
                            $"Error closing presentation during force quit: {ex.Message}"
                        );
                        // Continue with quit even if presentation close fails
                    }
                    _currentPresentation = null;
                }

                // Then quit PowerPoint with all alerts disabled
                if (_powerPointApp != null)
                {
                    try
                    {
                        // Disable all alerts
                        _powerPointApp.DisplayAlerts = 0; // ppAlertsNone = 0

                        // Quit the application
                        _powerPointApp.Quit();
                        Marshal.ReleaseComObject(_powerPointApp);
                        Console.WriteLine("PowerPoint quit successfully");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during PowerPoint quit: {ex.Message}");

                        // Try more aggressive approach
                        try
                        {
                            Marshal.FinalReleaseComObject(_powerPointApp);
                            Console.WriteLine("PowerPoint forcibly released");
                        }
                        catch (Exception finalEx)
                        {
                            Console.WriteLine($"Final release failed: {finalEx.Message}");
                            return false;
                        }
                    }
                    _powerPointApp = null;
                }

                _isInitialized = false;

                // Give some time for COM resources to be released
                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during force quit: {ex.Message}");
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
                if (disposing)
                {
                    // Clean up managed resources
                }

                // Clean up COM objects
                try
                {
                    // First, try to disable all alerts
                    if (_powerPointApp != null)
                    {
                        try
                        {
                            _powerPointApp.DisplayAlerts = 0; // ppAlertsNone = 0
                        }
                        catch
                        {
                            // Ignore errors when disabling alerts
                        }
                    }

                    // Handle presentation closure
                    if (_currentPresentation != null)
                    {
                        try
                        {
                            // Mark as saved to prevent save prompts
                            _currentPresentation.Saved = true;
                        }
                        catch
                        {
                            // Ignore setting saved state errors
                        }

                        ClosePresentation();

                        try
                        {
                            // Release COM object
                            Marshal.ReleaseComObject(_currentPresentation);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(
                                $"Error releasing presentation COM object: {ex.Message}"
                            );
                        }
                    }

                    if (_powerPointApp != null)
                    {
                        try
                        {
                            // Try the standard way first
                            _powerPointApp.Quit();
                            Marshal.ReleaseComObject(_powerPointApp);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(
                                $"Error during standard PowerPoint closure: {ex.Message}"
                            );

                            // If standard quit fails, try to force termination
                            try
                            {
                                // Last resort - just release the COM object without proper quit
                                Marshal.FinalReleaseComObject(_powerPointApp);
                            }
                            catch (Exception finalEx)
                            {
                                Console.WriteLine(
                                    $"Error during forced PowerPoint release: {finalEx.Message}"
                                );
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during disposal: {ex.Message}");
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

        public bool PreloadPresentation(string filePath)
        {
            try
            {
                // Quick validation
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    return false;
                }

                // Get file metadata for cache
                string fullPath = Path.GetFullPath(filePath);

                // Just store the path - simplify the caching mechanism
                _preloadedPresentations[fullPath] = fullPath;

                // Ensure PowerPoint is initialized
                if (!_isInitialized)
                {
                    InitializePowerPoint();
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in PreloadPresentation: {ex.Message}");
                return false;
            }
        }

        public bool ReopenLastPresentation(bool startSlideShow = true)
        {
            if (string.IsNullOrEmpty(_lastOpenedPath) || !File.Exists(_lastOpenedPath))
            {
                return false;
            }

            return OpenPresentation(_lastOpenedPath, startSlideShow, true);
        }

        // Add this helper method for retrying operations
        private T RetryOperation<T>(Func<T> operation, T defaultValue, int maxRetries = 2)
        {
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    return operation();
                }
                catch (Exception ex)
                {
                    if (attempt == maxRetries)
                    {
                        Console.WriteLine(
                            $"Operation failed after {maxRetries} attempts: {ex.Message}"
                        );
                    }
                    else
                    {
                        Console.WriteLine($"Attempt {attempt} failed: {ex.Message}. Retrying...");
                        // Short pause before retry
                        Thread.Sleep(100);
                    }
                }
            }
            return defaultValue;
        }

        // Add this helper method for more reliable COM property access
        private dynamic GetProperty(dynamic comObject, string propertyName)
        {
            try
            {
                if (comObject == null)
                    return null;

                // This uses reflection to get the property more reliably
                return comObject
                    .GetType()
                    .InvokeMember(
                        propertyName,
                        System.Reflection.BindingFlags.GetProperty,
                        null,
                        comObject,
                        null
                    );
            }
            catch
            {
                return null;
            }
        }
    }
}
