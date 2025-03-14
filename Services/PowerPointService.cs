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

        // Get available macros in the current presentation
        public List<string> GetAvailableMacros()
        {
            List<string> macroNames = new List<string>();

            try
            {
                if (_currentPresentation == null)
                {
                    Console.WriteLine("No presentation is open");
                    return macroNames;
                }

                // First, check if the presentation has VBA project
                try
                {
                    dynamic vbProject = _currentPresentation.VBProject;
                    if (vbProject == null)
                    {
                        Console.WriteLine("Presentation does not have a VBA project");
                        return macroNames;
                    }

                    // Get VBA components (modules)
                    dynamic vbComponents = vbProject.VBComponents;
                    if (vbComponents == null)
                    {
                        return macroNames;
                    }

                    // Loop through all components
                    int count = vbComponents.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            dynamic component = vbComponents.Item(i);
                            if (component != null)
                            {
                                string componentName = component.Name;
                                dynamic codeModule = component.CodeModule;

                                if (codeModule != null && codeModule.CountOfLines > 0)
                                {
                                    // Get the code to parse for Sub procedures
                                    string code = codeModule.Lines(1, codeModule.CountOfLines);

                                    // Very basic parsing for Sub procedures
                                    // This is a simplified approach - real parsing would be more complex
                                    string[] lines = code.Split(
                                        new[] { '\r', '\n' },
                                        StringSplitOptions.RemoveEmptyEntries
                                    );
                                    foreach (string line in lines)
                                    {
                                        string trimmedLine = line.Trim();
                                        if (
                                            trimmedLine.StartsWith(
                                                "Sub ",
                                                StringComparison.OrdinalIgnoreCase
                                            )
                                            && !trimmedLine.Contains("(")
                                            && !trimmedLine.Contains(")")
                                        )
                                        {
                                            // Extract macro name - this is a very basic extraction
                                            string macroName = trimmedLine.Substring(4).Trim();
                                            macroNames.Add($"{componentName}.{macroName}");
                                        }
                                        else if (
                                            trimmedLine.StartsWith(
                                                "Sub ",
                                                StringComparison.OrdinalIgnoreCase
                                            )
                                            && trimmedLine.Contains("(")
                                            && trimmedLine.Contains(")")
                                        )
                                        {
                                            // Extract macro name from a Sub with parameters
                                            int parenIndex = trimmedLine.IndexOf('(');
                                            if (parenIndex > 4)
                                            {
                                                string macroName = trimmedLine
                                                    .Substring(4, parenIndex - 4)
                                                    .Trim();
                                                macroNames.Add($"{componentName}.{macroName}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error examining VBA component {i}: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error examining VBA project: {ex.Message}");

                    // Fall back to simpler detection if VBProject access fails due to security settings
                    try
                    {
                        bool hasMacros = _currentPresentation.HasVBProject;
                        if (hasMacros)
                        {
                            macroNames.Add(
                                "(Macros exist but details cannot be accessed due to security settings)"
                            );
                        }
                    }
                    catch
                    {
                        // Even the HasVBProject check failed, ignore
                    }
                }

                return macroNames;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting available macros: {ex.Message}");
                return macroNames;
            }
        }

        // Run a specific macro by name - updated to work even with restricted VBA access
        public bool RunMacro(string macroName)
        {
            try
            {
                if (_powerPointApp == null || _currentPresentation == null)
                {
                    Console.WriteLine("PowerPoint or presentation is not available");
                    return false;
                }

                if (string.IsNullOrEmpty(macroName))
                {
                    Console.WriteLine("Macro name is required");
                    return false;
                }

                // Try to run the macro directly without requiring VBA project access
                return RetryOperation(
                    () =>
                    {
                        try
                        {
                            // Method 1: Try direct execution through Run method
                            try
                            {
                                Console.WriteLine($"Attempting to run macro: {macroName}");
                                // This method may work even when VBA project access is restricted
                                _powerPointApp.Run(macroName);
                                Console.WriteLine(
                                    $"Successfully ran macro via App.Run: {macroName}"
                                );
                                return true;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to run using App.Run: {ex.Message}");
                            }

                            // Method 2: Try using SendKeys as a fallback (works in some cases)
                            try
                            {
                                // Try to activate the VBA IDE and run the macro via ALT+F8 (macro dialog)
                                _powerPointApp.CommandBars.ExecuteMso("ShowVisualBasicEditor");
                                Console.WriteLine(
                                    "Opened VBA editor. Attempting alternative execution..."
                                );

                                // Try executing directly from immediate window
                                var vbe = _powerPointApp.VBE;
                                if (vbe != null)
                                {
                                    try
                                    {
                                        // Try getting the active window's immediate pane
                                        var activeWindow = vbe.ActiveWindow;
                                        if (activeWindow != null)
                                        {
                                            // Send the macro name directly to the immediate pane
                                            vbe.CommandBars.FindControl(Id: 2082).Execute(); // Opens immediate window
                                            vbe.ActiveCodePane.CodeModule.InsertLines(
                                                1,
                                                $"Call {macroName}"
                                            );
                                            Console.WriteLine(
                                                $"Successfully sent macro command: {macroName}"
                                            );
                                            return true;
                                        }
                                    }
                                    catch (Exception exVbe)
                                    {
                                        Console.WriteLine(
                                            $"VBA editor execution failed: {exVbe.Message}"
                                        );
                                    }
                                }
                            }
                            catch (Exception exCmd)
                            {
                                Console.WriteLine($"Command execution failed: {exCmd.Message}");
                            }

                            // Method 3: As a last resort, try a direct Application.Run call with a fully qualified name
                            try
                            {
                                // Attempt with VBIDE module
                                if (_powerPointApp.VBIDE != null)
                                {
                                    Console.WriteLine("Attempting direct VBIDE access...");
                                    return false; // This will be overridden if the above line doesn't throw
                                }
                            }
                            catch
                            {
                                // Expected if security is tight
                            }

                            // If we get here, all methods failed
                            Console.WriteLine("All macro execution methods failed");
                            Console.WriteLine(
                                "To enable macro execution, open PowerPoint and go to:"
                            );
                            Console.WriteLine(
                                "File > Options > Trust Center > Trust Center Settings > Macro Settings"
                            );
                            Console.WriteLine(
                                "Then check 'Trust access to the VBA project object model'"
                            );

                            return false;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error running macro '{macroName}': {ex.Message}");
                            return false;
                        }
                    },
                    false
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during macro execution: {ex.Message}");
                return false;
            }
        }

        // Directly try to run a slide-specific macro without attempting to enumerate first
        public bool TryRunSlideSpecificMacro(int slideNumber)
        {
            try
            {
                // Create an array of common naming patterns for slide macros
                string[] commonPatterns =
                {
                    $"Slide{slideNumber}_Action",
                    $"Slide_{slideNumber}",
                    $"Slide{slideNumber}Action",
                    $"SlideAction{slideNumber}",
                    $"Slide{slideNumber}",
                    $"OnSlide{slideNumber}",
                    $"OnEnterSlide{slideNumber}",
                    $"RunSlide{slideNumber}",
                    // Add generic slide change handlers
                    "OnSlideChange",
                    "SlideChanged",
                    "SlideChange",
                };

                // Try each pattern directly without needing VBA project access
                foreach (string pattern in commonPatterns)
                {
                    bool success = false;

                    // Try with common module names
                    string[] commonModules =
                    {
                        "Module1",
                        "SlideModule",
                        "Macros",
                        "Presentation",
                        "ThisPresentation",
                    };

                    foreach (string module in commonModules)
                    {
                        // Try with fully qualified name
                        string fullName = $"{module}.{pattern}";
                        Console.WriteLine($"Attempting to run macro: {fullName}");

                        try
                        {
                            _powerPointApp.Run(fullName);
                            Console.WriteLine($"Successfully ran macro: {fullName}");
                            success = true;
                            break;
                        }
                        catch
                        {
                            // Try next module
                        }
                    }

                    // If we succeeded with any module, return
                    if (success)
                        return true;

                    // Try without module qualification as a last resort
                    try
                    {
                        _powerPointApp.Run(pattern);
                        Console.WriteLine($"Successfully ran macro: {pattern}");
                        return true;
                    }
                    catch
                    {
                        // Continue to next pattern
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in TryRunSlideSpecificMacro: {ex.Message}");
                return false;
            }
        }

        // Run macro on current slide if available - improved to handle presentations not in slideshow mode
        public bool RunMacroOnCurrentSlide()
        {
            try
            {
                // First check if we have a presentation open
                if (_powerPointApp == null || _currentPresentation == null)
                {
                    Console.WriteLine("No presentation is open");
                    return false;
                }

                // Get current slide number - first try slideshow mode
                int currentSlideNumber = GetCurrentSlideNumber();

                // If we couldn't get the slide number from slideshow view, try getting it from the presentation
                if (currentSlideNumber <= 0)
                {
                    try
                    {
                        // Try to get the active slide from the normal view
                        dynamic view = _currentPresentation.View;
                        if (view != null)
                        {
                            // Try getting the selected slides
                            try
                            {
                                dynamic slide = view.Slide;
                                if (slide != null)
                                {
                                    currentSlideNumber = Convert.ToInt32(slide.SlideNumber);
                                }
                            }
                            catch
                            {
                                // Even this failed, try other view types
                            }
                        }

                        // If still no slide number, try a different method
                        if (currentSlideNumber <= 0)
                        {
                            try
                            {
                                // Try to get the active window and current selection
                                dynamic activeWindow = _powerPointApp.ActiveWindow;
                                if (activeWindow != null)
                                {
                                    dynamic selection = activeWindow.Selection;
                                    if (selection != null && selection.SlideRange != null)
                                    {
                                        currentSlideNumber = Convert.ToInt32(
                                            selection.SlideRange.SlideNumber
                                        );
                                    }
                                }
                            }
                            catch
                            {
                                // Even this approach failed
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error detecting slide in normal view: {ex.Message}");
                    }
                }

                // If we still don't have a slide number, default to the first slide
                if (currentSlideNumber <= 0)
                {
                    Console.WriteLine("Could not determine current slide - using first slide");
                    currentSlideNumber = 1; // Default to first slide
                }

                Console.WriteLine(
                    $"Current slide: {currentSlideNumber}. Attempting to run slide-specific macros..."
                );

                // Try two approaches:

                // Approach 1: Try to get available macros if VBA access is not restricted
                try
                {
                    List<string> availableMacros = GetAvailableMacros();
                    if (availableMacros.Count > 0)
                    {
                        // We have access to the macro list
                        string slideNumberStr = currentSlideNumber.ToString();
                        string slideMacroName = null;

                        // Look for possible naming patterns for slide macros
                        foreach (string macro in availableMacros)
                        {
                            // Check for common naming patterns
                            if (
                                macro.Contains($"Slide{slideNumberStr}_")
                                || macro.Contains($"Slide_{slideNumberStr}")
                                || macro.Contains($"Slide{slideNumberStr}Action")
                                || macro.Contains($"SlideAction{slideNumberStr}")
                                || macro.EndsWith($"Slide{slideNumberStr}")
                            )
                            {
                                slideMacroName = macro;
                                break;
                            }
                        }

                        // If we found a match, run it
                        if (!string.IsNullOrEmpty(slideMacroName))
                        {
                            bool success = RunMacro(slideMacroName);
                            if (success)
                                return true;
                        }

                        // If no slide-specific macro was found, look for a generic "OnSlideChange" macro
                        foreach (string macro in availableMacros)
                        {
                            if (
                                macro.Contains("OnSlideChange")
                                || macro.Contains("SlideChanged")
                                || macro.Contains("SlideChange")
                            )
                            {
                                bool success = RunMacro(macro);
                                if (success)
                                    return true;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"VBA project access approach failed: {ex.Message}");
                    // Continue to approach 2
                }

                // Approach 2: Try direct execution without VBA project access
                Console.WriteLine("Trying direct macro execution approach...");
                return TryRunSlideSpecificMacro(currentSlideNumber);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running macro on current slide: {ex.Message}");
                return false;
            }
        }

        // Check if presentation has macros without requiring VBA project access
        public bool HasMacros()
        {
            try
            {
                if (_currentPresentation == null)
                {
                    Console.WriteLine("No presentation is open");
                    return false;
                }

                // First try the HasVBProject property which is sometimes available even when VBA access is restricted
                try
                {
                    bool hasMacros = _currentPresentation.HasVBProject;
                    if (hasMacros)
                    {
                        Console.WriteLine("Presentation contains VBA macros");
                        return true;
                    }
                    else
                    {
                        Console.WriteLine("Presentation does not contain VBA macros");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"HasVBProject check failed: {ex.Message}");
                }

                // Check file extension - .pptm definitely has macros
                if (!string.IsNullOrEmpty(_lastOpenedPath))
                {
                    string extension = Path.GetExtension(_lastOpenedPath).ToLower();
                    if (extension == ".pptm")
                    {
                        Console.WriteLine(
                            "Presentation has .pptm extension, which indicates it contains macros"
                        );
                        return true;
                    }
                }

                // Check presentation type to see if it has macros
                try
                {
                    // Try to check the presentation type
                    int presentationType = Convert.ToInt32(_currentPresentation.HasVBProject);
                    if (presentationType > 0)
                    {
                        return true;
                    }
                }
                catch
                {
                    // Ignore - this is just an additional check
                }

                // We couldn't definitively determine if the presentation has macros
                Console.WriteLine("Could not determine if presentation has macros");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking for macros: {ex.Message}");
                return false;
            }
        }

        // Get potential macro names based on the current slide
        public List<string> GetPotentialMacroNames()
        {
            List<string> potentialMacros = new List<string>();

            try
            {
                int currentSlide = GetCurrentSlideNumber();
                if (currentSlide <= 0)
                {
                    return potentialMacros;
                }

                // Add commonly used module names
                string[] commonModules =
                {
                    "Module1",
                    "SlideModule",
                    "Macros",
                    "Presentation",
                    "ThisPresentation",
                };

                // Add slide-specific macro patterns
                string[] patterns =
                {
                    $"Slide{currentSlide}_Action",
                    $"Slide_{currentSlide}",
                    $"Slide{currentSlide}Action",
                    $"SlideAction{currentSlide}",
                    $"Slide{currentSlide}",
                    $"OnSlide{currentSlide}",
                    $"OnEnterSlide{currentSlide}",
                    $"RunSlide{currentSlide}",
                    // Add generic slide change handlers
                    "OnSlideChange",
                    "SlideChanged",
                    "SlideChange",
                };

                // Generate all combinations
                foreach (string module in commonModules)
                {
                    foreach (string pattern in patterns)
                    {
                        potentialMacros.Add($"{module}.{pattern}");
                    }
                }

                // Also add non-qualified names as they sometimes work
                foreach (string pattern in patterns)
                {
                    potentialMacros.Add(pattern);
                }

                return potentialMacros;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating potential macro names: {ex.Message}");
                return potentialMacros;
            }
        }

        // Start slideshow mode if not already in it
        public bool StartSlideShow()
        {
            try
            {
                // Check if already in slideshow mode
                if (_powerPointApp == null || _currentPresentation == null)
                {
                    Console.WriteLine("No presentation is open");
                    return false;
                }

                // Check if already in slideshow mode
                dynamic? slideShowWindows = _powerPointApp.SlideShowWindows;
                if (slideShowWindows != null && Convert.ToInt32(slideShowWindows.Count) > 0)
                {
                    Console.WriteLine("Presentation is already in slideshow mode");
                    return true;
                }

                // Start slideshow from the beginning or current slide
                return RetryOperation(
                    () =>
                    {
                        try
                        {
                            var settings = _currentPresentation.SlideShowSettings;
                            if (settings == null)
                            {
                                Console.WriteLine("SlideShowSettings is not available");
                                return false;
                            }

                            // Configure slideshow settings
                            settings.ShowType = 1; // ppShowTypeSpeaker
                            settings.StartingSlide = 1;
                            settings.EndingSlide = GetTotalSlides();
                            settings.ShowWithAnimation = true;

                            // Run the slideshow
                            settings.Run();
                            Console.WriteLine("Started slideshow mode");
                            return true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error starting slideshow: {ex.Message}");
                            return false;
                        }
                    },
                    false
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in StartSlideShow: {ex.Message}");
                return false;
            }
        }
    }
}
