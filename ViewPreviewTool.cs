using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace ViewPreviewTool
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class ViewPreviewApplication : IExternalApplication
    {
        private static ViewPreviewForm previewForm;
        private static UIApplication uiApp;
        private static bool isMonitoring = false;
        private static ElementId lastViewId = null;
        private static System.Windows.Forms.Timer positionTimer;
        
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create ribbon panel
                RibbonPanel ribbonPanel = application.CreateRibbonPanel("BIM Ops Studio");
                
                // Create toggle button
                PushButtonData buttonData = new PushButtonData(
                    "ViewPreviewTool",
                    "View\nPreview",
                    System.Reflection.Assembly.GetExecutingAssembly().Location,
                    "ViewPreviewTool.TogglePreviewCommand");
                
                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;
                pushButton.ToolTip = "Toggle view preview window";
                pushButton.LongDescription = "Click to open preview window. Then select views in Project Browser to see previews.";
                
                // Try to set icon
                try
                {
                    pushButton.LargeImage = CreateButtonImage(32);
                    pushButton.Image = CreateButtonImage(16);
                }
                catch { }
                
                // Subscribe to Idling event
                application.Idling += OnIdling;
                
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
        
        public Result OnShutdown(UIControlledApplication application)
        {
            application.Idling -= OnIdling;
            
            if (positionTimer != null)
            {
                positionTimer.Stop();
                positionTimer.Dispose();
            }
            
            if (previewForm != null && !previewForm.IsDisposed)
            {
                previewForm.Close();
                previewForm.Dispose();
            }
            return Result.Succeeded;
        }
        
        private void OnIdling(object sender, IdlingEventArgs e)
        {
            if (!isMonitoring || uiApp == null) return;
            
            try
            {
                UIDocument uidoc = uiApp.ActiveUIDocument;
                if (uidoc == null) return;
                
                // Check current selection
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count != 1) return;
                
                ElementId selectedId = selectedIds.First();
                
                // Check if selection changed
                if (lastViewId != null && lastViewId.IntegerValue == selectedId.IntegerValue) return;
                
                // Get element
                Element elem = uidoc.Document.GetElement(selectedId);
                Autodesk.Revit.DB.View view = elem as Autodesk.Revit.DB.View;
                
                if (view != null && view.CanBePrinted)
                {
                    lastViewId = selectedId;
                    
                    // Update preview
                    if (previewForm != null && !previewForm.IsDisposed)
                    {
                        previewForm.UpdatePreview(uidoc.Document, view);
                    }
                }
            }
            catch { }
        }
        
        public static void StartMonitoring(UIApplication app)
        {
            uiApp = app;
            isMonitoring = true;
            
            if (previewForm == null || previewForm.IsDisposed)
            {
                previewForm = new ViewPreviewForm();
                PositionPreviewWindow();
                
                // Set up timer to reposition window if Project Browser moves
                positionTimer = new System.Windows.Forms.Timer();
                positionTimer.Interval = 500; // Check every 500ms
                positionTimer.Tick += (s, e) => PositionPreviewWindow();
                positionTimer.Start();
            }
            
            previewForm.Show();
            previewForm.BringToFront();
            
            // Check if there's already a view selected
            try
            {
                UIDocument uidoc = app.ActiveUIDocument;
                if (uidoc != null)
                {
                    ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                    if (selectedIds.Count == 1)
                    {
                        Element elem = uidoc.Document.GetElement(selectedIds.First());
                        Autodesk.Revit.DB.View view = elem as Autodesk.Revit.DB.View;
                        if (view != null && view.CanBePrinted)
                        {
                            lastViewId = selectedIds.First();
                            previewForm.UpdatePreview(uidoc.Document, view);
                        }
                    }
                }
            }
            catch { }
        }
        
        public static void StopMonitoring()
        {
            isMonitoring = false;
            lastViewId = null;
            
            if (positionTimer != null)
            {
                positionTimer.Stop();
                positionTimer.Dispose();
                positionTimer = null;
            }
            
            if (previewForm != null && !previewForm.IsDisposed)
            {
                previewForm.Close();
                previewForm.Dispose();
                previewForm = null;
            }
        }
        
        private static void PositionPreviewWindow()
        {
            if (previewForm == null || previewForm.IsDisposed) return;
            
            previewForm.StartPosition = FormStartPosition.Manual;
            
            try
            {
                // Get the main Revit window handle
                IntPtr revitHandle = Process.GetCurrentProcess().MainWindowHandle;
                if (revitHandle == IntPtr.Zero) return;
                
                // Find the Project Browser window
                IntPtr projectBrowserHandle = FindProjectBrowserWindow(revitHandle);
                
                if (projectBrowserHandle != IntPtr.Zero)
                {
                    // Get Project Browser position
                    RECT browserRect;
                    if (GetWindowRect(projectBrowserHandle, out browserRect))
                    {
                        // Position our window to the left of Project Browser
                        int gap = 10; // Gap between windows
                        int x = browserRect.Left - previewForm.Width - gap;
                        int y = browserRect.Top;
                        
                        // Fine-tune position to fit within canvas
                        x -= 5; // Move 5 pixels to the left
                        y += 3; // Move 3 pixels down
                        
                        // Get the screen that contains the Project Browser
                        var browserBounds = new System.Drawing.Rectangle(
                            browserRect.Left, browserRect.Top,
                            browserRect.Right - browserRect.Left,
                            browserRect.Bottom - browserRect.Top);
                        
                        var screen = Screen.FromRectangle(browserBounds);
                        
                        // If there's not enough space on the left, try other positions
                        if (x < screen.WorkingArea.Left)
                        {
                            // Try to position on the right side
                            x = browserRect.Right + gap;
                            x -= 5; // Apply same offset
                            
                            // If that's also off-screen, position inside the Revit window
                            if (x + previewForm.Width > screen.WorkingArea.Right)
                            {
                                // Get Revit window bounds
                                RECT revitRect;
                                if (GetWindowRect(revitHandle, out revitRect))
                                {
                                    // Position in the drawing area (left of Project Browser, inside Revit)
                                    x = browserRect.Left - previewForm.Width - gap;
                                    x -= 5; // Apply same offset
                                    if (x < revitRect.Left + 50) // Leave some margin
                                    {
                                        x = revitRect.Left + 50;
                                    }
                                }
                            }
                        }
                        
                        // Ensure y is within screen bounds
                        if (y < screen.WorkingArea.Top)
                        {
                            y = screen.WorkingArea.Top;
                        }
                        
                        // Ensure window is fully visible
                        if (y + previewForm.Height > screen.WorkingArea.Bottom)
                        {
                            y = screen.WorkingArea.Bottom - previewForm.Height;
                        }
                        
                        previewForm.Location = new System.Drawing.Point(x, y);
                        return;
                    }
                }
                
                // Fallback: If we can't find Project Browser, use smart positioning
                RECT revitRectFallback;
                if (GetWindowRect(revitHandle, out revitRectFallback))
                {
                    var revitBounds = new System.Drawing.Rectangle(
                        revitRectFallback.Left, revitRectFallback.Top,
                        revitRectFallback.Right - revitRectFallback.Left,
                        revitRectFallback.Bottom - revitRectFallback.Top);
                    
                    var screen = Screen.FromRectangle(revitBounds);
                    
                    // Position on the right side of Revit window (where Project Browser usually is)
                    // but leave space for a typical Project Browser width
                    int typicalBrowserWidth = 400;
                    int x = revitRectFallback.Right - typicalBrowserWidth - previewForm.Width - 20;
                    int y = revitRectFallback.Top + 150;
                    
                    // Fine-tune position to fit within canvas
                    x -= 5; // Move 5 pixels to the left
                    y += 3; // Move 3 pixels down
                    
                    // Ensure it's within the Revit window
                    if (x < revitRectFallback.Left + 50)
                    {
                        x = revitRectFallback.Left + 50;
                    }
                    
                    previewForm.Location = new System.Drawing.Point(x, y);
                }
            }
            catch { }
        }
        
        private static IntPtr FindProjectBrowserWindow(IntPtr revitHandle)
        {
            IntPtr foundHandle = IntPtr.Zero;
            
            // Enumerate all child windows of Revit
            EnumChildWindows(revitHandle, (hwnd, lParam) =>
            {
                // Get window class name
                StringBuilder className = new StringBuilder(256);
                GetClassName(hwnd, className, className.Capacity);
                
                // Get window text
                StringBuilder windowText = new StringBuilder(256);
                GetWindowText(hwnd, windowText, windowText.Capacity);
                
                string classStr = className.ToString();
                string textStr = windowText.ToString();
                
                // Project Browser detection logic
                // Look for windows that might be the Project Browser
                if (textStr.Contains("Project Browser") || 
                    textStr.Contains("project browser") ||
                    (classStr.Contains("AfxWnd") && IsProjectBrowserLikeWindow(hwnd)))
                {
                    foundHandle = hwnd;
                    return false; // Stop enumeration
                }
                
                return true; // Continue enumeration
            }, IntPtr.Zero);
            
            // If direct search failed, try deeper search
            if (foundHandle == IntPtr.Zero)
            {
                foundHandle = DeepSearchForProjectBrowser(revitHandle);
            }
            
            return foundHandle;
        }
        
        private static IntPtr DeepSearchForProjectBrowser(IntPtr parent)
        {
            IntPtr result = IntPtr.Zero;
            
            EnumChildWindows(parent, (hwnd, lParam) =>
            {
                // Check if this could be a docking panel
                RECT rect;
                if (GetWindowRect(hwnd, out rect))
                {
                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;
                    
                    // Project Browser is typically 300-500 pixels wide and tall
                    if (width > 250 && width < 600 && height > 400)
                    {
                        // Check if it's on the right side of the screen
                        var screen = Screen.FromHandle(hwnd);
                        if (rect.Right > screen.WorkingArea.Right - 600)
                        {
                            // This could be the Project Browser
                            result = hwnd;
                            return false;
                        }
                    }
                }
                
                // Recursively search children
                IntPtr childResult = DeepSearchForProjectBrowser(hwnd);
                if (childResult != IntPtr.Zero)
                {
                    result = childResult;
                    return false;
                }
                
                return true;
            }, IntPtr.Zero);
            
            return result;
        }
        
        private static bool IsProjectBrowserLikeWindow(IntPtr hwnd)
        {
            // Check window properties to see if it could be the Project Browser
            RECT rect;
            if (GetWindowRect(hwnd, out rect))
            {
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                
                // Project Browser is typically a tall, narrow window on the right
                return width > 250 && width < 600 && height > 400;
            }
            
            return false;
        }
        
        private System.Windows.Media.Imaging.BitmapSource CreateButtonImage(int size)
        {
            try
            {
                // Create a simple blue square icon
                Bitmap bitmap = new Bitmap(size, size);
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(System.Drawing.Color.Transparent);
                    using (Brush brush = new SolidBrush(System.Drawing.Color.FromArgb(0, 122, 204)))
                    {
                        g.FillRectangle(brush, 2, 2, size - 4, size - 4);
                    }
                    g.DrawRectangle(Pens.DarkBlue, 2, 2, size - 5, size - 5);
                    
                    // Add "VP" text
                    using (Font font = new Font("Arial", size / 3, FontStyle.Bold))
                    using (Brush textBrush = new SolidBrush(System.Drawing.Color.White))
                    {
                        string text = "VP";
                        SizeF textSize = g.MeasureString(text, font);
                        float x = (size - textSize.Width) / 2;
                        float y = (size - textSize.Height) / 2;
                        g.DrawString(text, font, textBrush, x, y);
                    }
                }
                
                IntPtr hBitmap = bitmap.GetHbitmap();
                bitmap.Dispose();
                try
                {
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                        hBitmap,
                        IntPtr.Zero,
                        System.Windows.Int32Rect.Empty,
                        System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());
                }
                finally
                {
                    DeleteObject(hBitmap);
                }
            }
            catch
            {
                return null;
            }
        }
        
        #region P/Invoke declarations
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
        
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);
        
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        
        private const uint GW_CHILD = 5;
        private const uint GW_HWNDNEXT = 2;
        
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion
    }
    
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class TogglePreviewCommand : IExternalCommand
    {
        private static bool isActive = false;
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!isActive)
                {
                    // Start monitoring
                    ViewPreviewApplication.StartMonitoring(commandData.Application);
                    isActive = true;
                }
                else
                {
                    // Stop monitoring
                    ViewPreviewApplication.StopMonitoring();
                    isActive = false;
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
    
    public class ViewPreviewForm : System.Windows.Forms.Form
    {
        private Document doc;
        private Autodesk.Revit.DB.View view;
        private PictureBox pictureBox;
        private System.Windows.Forms.Panel containerPanel;
        private System.Windows.Forms.Label statusLabel;
        private bool isDragging = false;
        private System.Drawing.Point dragStartPoint;
        private float currentZoomFactor = 1.0f;
        private Image originalImage = null;
        private bool isPanning = false;
        private System.Drawing.Point panStartPoint;
        private System.Drawing.Point scrollStartPoint;
        
        public ViewPreviewForm()
        {
            InitializeForm();
        }
        
        private void InitializeForm()
        {
            this.Text = "View Preview";
            this.Size = new Size(850, 700);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(600, 500);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            
            // Header panel
            System.Windows.Forms.Panel headerPanel = new System.Windows.Forms.Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 80;
            headerPanel.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            
            // Add drag functionality to header
            headerPanel.MouseDown += HeaderPanel_MouseDown;
            headerPanel.MouseMove += HeaderPanel_MouseMove;
            headerPanel.MouseUp += HeaderPanel_MouseUp;
            
            // Title label - FIXED with AutoEllipsis
            System.Windows.Forms.Label titleLabel = new System.Windows.Forms.Label();
            titleLabel.Name = "titleLabel";
            titleLabel.Font = new Font("Segoe UI", 18, FontStyle.Bold);
            titleLabel.ForeColor = System.Drawing.Color.White;
            titleLabel.Dock = DockStyle.None;
            titleLabel.AutoSize = false;
            titleLabel.Location = new System.Drawing.Point(10, 10);  // Add 10px padding from left
            titleLabel.Size = new Size(this.ClientSize.Width - 20, 40);  // Subtract padding from both sides
            titleLabel.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            titleLabel.BackColor = System.Drawing.Color.Transparent;
            titleLabel.AutoEllipsis = true;  // This will add "..." when text is too long
            titleLabel.UseMnemonic = false;
            titleLabel.Text = "Select a view in Project Browser";
            titleLabel.TextAlign = ContentAlignment.MiddleCenter;
            titleLabel.MouseDown += HeaderPanel_MouseDown;
            titleLabel.MouseMove += HeaderPanel_MouseMove;
            titleLabel.MouseUp += HeaderPanel_MouseUp;
            // Type label
            System.Windows.Forms.Label typeLabel = new System.Windows.Forms.Label();
            typeLabel.Name = "typeLabel";
            typeLabel.Font = new Font("Segoe UI", 11);
            typeLabel.ForeColor = System.Drawing.Color.FromArgb(220, 220, 220);
            typeLabel.Dock = DockStyle.None;
            typeLabel.AutoSize = false;
            typeLabel.Location = new System.Drawing.Point(10, 48);  // Add padding to match title
            typeLabel.Size = new Size(this.ClientSize.Width - 20, 25);  // Match title padding
            typeLabel.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            typeLabel.BackColor = System.Drawing.Color.Transparent;
            typeLabel.AutoEllipsis = true;  // Also add ellipsis for type label
            typeLabel.Text = "Click on views to see preview";
            typeLabel.TextAlign = ContentAlignment.MiddleCenter;
            typeLabel.MouseDown += HeaderPanel_MouseDown;
            typeLabel.MouseMove += HeaderPanel_MouseMove;
            typeLabel.MouseUp += HeaderPanel_MouseUp;
            
            // Status panel
            System.Windows.Forms.Panel statusPanel = new System.Windows.Forms.Panel();
            statusPanel.Dock = DockStyle.Bottom;
            statusPanel.Height = 30;
            statusPanel.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            
            statusLabel = new System.Windows.Forms.Label();
            statusLabel.Text = "Ready - Select a view to preview";
            statusLabel.Dock = DockStyle.Fill;
            statusLabel.TextAlign = ContentAlignment.MiddleCenter;
            statusLabel.Font = new Font("Segoe UI", 9);
            statusLabel.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80);
            statusPanel.Controls.Add(statusLabel);
            
            // Picture box container
            containerPanel = new System.Windows.Forms.Panel();
            containerPanel.Dock = DockStyle.Fill;
            containerPanel.BackColor = System.Drawing.Color.White;
            containerPanel.Padding = new Padding(10);
            containerPanel.AutoScroll = false;
            
            // Picture box - DO NOT use Dock.Fill for zoom/pan to work
            pictureBox = new PictureBox();
            pictureBox.Location = new System.Drawing.Point(0, 0);
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox.BackColor = System.Drawing.Color.White;
            
            containerPanel.Controls.Add(pictureBox);
            
            // Add controls
            this.Controls.Add(containerPanel);
            this.Controls.Add(statusPanel);
            this.Controls.Add(headerPanel);
            
            // Add labels to header panel after form is set up
            headerPanel.Controls.Add(titleLabel);
            headerPanel.Controls.Add(typeLabel);
            
            // Add zoom functionality on double-click
            pictureBox.DoubleClick += (s, e) =>
            {
                if (pictureBox.Image == null || originalImage == null) return;
                
                // Reset zoom and fit to window
                currentZoomFactor = 1.0f;
                
                // Dispose zoomed image if not the original
                if (pictureBox.Image != originalImage)
                {
                    pictureBox.Image.Dispose();
                }
                
                pictureBox.Image = originalImage;
                pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                containerPanel.AutoScroll = false;
                statusLabel.Text = "Preview loaded - Scroll to zoom, double-click to fit";
            };
            
            // Add mouse wheel zoom functionality with proper cursor-centered zoom
            pictureBox.MouseWheel += (s, e) =>
            {
                if (pictureBox.Image == null || originalImage == null) return;
                
                // Get mouse position relative to the picture box
                System.Drawing.Point mousePos = pictureBox.PointToClient(System.Windows.Forms.Control.MousePosition);
                
                // Store the current scroll position before zoom
                System.Drawing.Point oldScroll = new System.Drawing.Point(
                    Math.Abs(containerPanel.AutoScrollPosition.X),
                    Math.Abs(containerPanel.AutoScrollPosition.Y)
                );
                
                // Calculate the mouse position in the original image coordinates
                float mouseImageX = (mousePos.X + oldScroll.X) / currentZoomFactor;
                float mouseImageY = (mousePos.Y + oldScroll.Y) / currentZoomFactor;
                
                // Calculate new zoom factor
                float zoomSpeed = 0.1f; // Zoom speed
                float scaleFactor = e.Delta > 0 ? (1.0f + zoomSpeed) : (1.0f / (1.0f + zoomSpeed));
                float newZoomFactor = currentZoomFactor * scaleFactor;
                
                // Limit zoom range
                if (newZoomFactor < 0.1f) newZoomFactor = 0.1f;
                if (newZoomFactor > 10.0f) newZoomFactor = 10.0f;
                
                // Only update if zoom changed significantly
                if (Math.Abs(newZoomFactor - currentZoomFactor) < 0.001f) return;
                
                float oldZoom = currentZoomFactor;
                currentZoomFactor = newZoomFactor;
                
                // Switch to AutoSize mode if needed
                if (pictureBox.SizeMode == PictureBoxSizeMode.Zoom)
                {
                    pictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
                    containerPanel.AutoScroll = true;
                }
                
                // Dispose old zoomed image
                if (pictureBox.Image != originalImage && pictureBox.Image != null)
                {
                    pictureBox.Image.Dispose();
                }
                
                // Create new zoomed image
                int newWidth = (int)(originalImage.Width * currentZoomFactor);
                int newHeight = (int)(originalImage.Height * currentZoomFactor);
                
                Bitmap zoomedImage = new Bitmap(newWidth, newHeight);
                using (Graphics g = Graphics.FromImage(zoomedImage))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                }
                
                pictureBox.Image = zoomedImage;
                
                // Calculate new scroll position to keep the mouse over the same point in the image
                int newScrollX = (int)(mouseImageX * currentZoomFactor - mousePos.X);
                int newScrollY = (int)(mouseImageY * currentZoomFactor - mousePos.Y);
                
                // Make sure scroll values are within bounds
                newScrollX = Math.Max(0, Math.Min(newScrollX, containerPanel.HorizontalScroll.Maximum));
                newScrollY = Math.Max(0, Math.Min(newScrollY, containerPanel.VerticalScroll.Maximum));
                
                // Apply the new scroll position
                containerPanel.AutoScrollPosition = new System.Drawing.Point(-newScrollX, -newScrollY);
                
                // Update status
                int zoomPercent = (int)(currentZoomFactor * 100);
                statusLabel.Text = $"Zoom: {zoomPercent}% - Scroll to zoom, double-click to fit, drag to pan";
            };
            
            // Add pan functionality to picture box - simplified and working
            pictureBox.MouseDown += (s, e) =>
            {
                // Allow panning with any mouse button when zoomed
                if (currentZoomFactor > 1.0f && containerPanel.AutoScroll)
                {
                    isPanning = true;
                    panStartPoint = e.Location;
                    // Get current scroll position
                    scrollStartPoint = new System.Drawing.Point(
                        Math.Abs(containerPanel.AutoScrollPosition.X),
                        Math.Abs(containerPanel.AutoScrollPosition.Y)
                    );
                    pictureBox.Cursor = Cursors.SizeAll;
                }
            };
            
            pictureBox.MouseMove += (s, e) =>
            {
                if (isPanning)
                {
                    // Calculate how much the mouse moved
                    int deltaX = e.X - panStartPoint.X;
                    int deltaY = e.Y - panStartPoint.Y;
                    
                    // Calculate new scroll position (subtract delta for natural panning)
                    int newScrollX = scrollStartPoint.X - deltaX;
                    int newScrollY = scrollStartPoint.Y - deltaY;
                    
                    // Ensure values are within bounds
                    newScrollX = Math.Max(0, newScrollX);
                    newScrollY = Math.Max(0, newScrollY);
                    
                    // Apply the scroll (AutoScrollPosition needs negative values)
                    containerPanel.AutoScrollPosition = new System.Drawing.Point(-newScrollX, -newScrollY);
                }
                else if (currentZoomFactor > 1.0f && containerPanel.AutoScroll)
                {
                    // Show hand cursor when pan is available
                    pictureBox.Cursor = Cursors.Hand;
                }
                else
                {
                    pictureBox.Cursor = Cursors.Default;
                }
            };
            
            pictureBox.MouseUp += (s, e) =>
            {
                // Stop panning on any mouse button release
                isPanning = false;
                
                // Update cursor
                if (currentZoomFactor > 1.0f && containerPanel.AutoScroll)
                {
                    pictureBox.Cursor = Cursors.Hand;
                }
                else
                {
                    pictureBox.Cursor = Cursors.Default;
                }
            };
            
            // Handle form resize to ensure labels stay full width with padding
            this.Resize += (s, e) =>
            {
                var titleLbl = this.Controls.Find("titleLabel", true).FirstOrDefault() as System.Windows.Forms.Label;
                var typeLbl = this.Controls.Find("typeLabel", true).FirstOrDefault() as System.Windows.Forms.Label;
                
                if (titleLbl != null && typeLbl != null)
                {
                    // Ensure labels span full width of header with padding
                    titleLbl.Width = this.ClientSize.Width - 20;  // 10px padding on each side
                    typeLbl.Width = this.ClientSize.Width - 20;
                }
            };
        }
        
        private void HeaderPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                dragStartPoint = new System.Drawing.Point(e.X, e.Y);
            }
        }
        
        private void HeaderPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                System.Drawing.Point currentScreenPos = PointToScreen(e.Location);
                Location = new System.Drawing.Point(
                    currentScreenPos.X - dragStartPoint.X,
                    currentScreenPos.Y - dragStartPoint.Y);
            }
        }
        
        private void HeaderPanel_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }
        
        public void UpdatePreview(Document document, Autodesk.Revit.DB.View v)
        {
            doc = document;
            view = v;
            
            // Reset zoom when changing views
            currentZoomFactor = 1.0f;
            pictureBox.Dock = DockStyle.Fill;
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            containerPanel.AutoScroll = false;
            
            // Update title
            var titleLabel = this.Controls.Find("titleLabel", true).FirstOrDefault() as System.Windows.Forms.Label;
            if (titleLabel != null)
            {
                titleLabel.Text = view.Name;
                // Add tooltip to show full text when truncated
                ToolTip titleTooltip = new ToolTip();
                titleTooltip.SetToolTip(titleLabel, view.Name);
            }
            
            // Update type
            var typeLabel = this.Controls.Find("typeLabel", true).FirstOrDefault() as System.Windows.Forms.Label;
            if (typeLabel != null)
            {
                typeLabel.Text = $"Type: {view.ViewType}";
            }
            
            // Update window title
            this.Text = $"View Preview: {view.Name}";
            
            // Clear existing image
            if (pictureBox.Image != null)
            {
                pictureBox.Image = null;
            }
            
            if (originalImage != null)
            {
                originalImage.Dispose();
                originalImage = null;
            }
            
            // Check if this is a schedule view
            if (view.ViewType == ViewType.Schedule || 
                view.ViewType == ViewType.ColumnSchedule || 
                view.ViewType == ViewType.PanelSchedule)
            {
                ShowScheduleMessage();
            }
            else
            {
                // Load new image immediately
                LoadViewImage();
            }
        }
        
        private void LoadViewImage()
        {
            try
            {
                statusLabel.Text = "Generating preview...";
                Application.DoEvents();
                
                string tempPath = Path.GetTempPath();
                string fileName = $"ViewPreview_{view.Id.IntegerValue}_{DateTime.Now.Ticks}";
                string filePath = Path.Combine(tempPath, fileName);
                
                var options = new ImageExportOptions
                {
                    FilePath = filePath,
                    FitDirection = FitDirectionType.Horizontal,
                    HLRandWFViewsFileType = ImageFileType.PNG,
                    ImageResolution = ImageResolution.DPI_150,
                    PixelSize = 1200,
                    ExportRange = ExportRange.SetOfViews
                };
                
                options.SetViewsAndSheets(new List<ElementId> { view.Id });
                
                bool exported = false;
                try
                {
                    doc.ExportImage(options);
                    exported = true;
                }
                catch (Exception ex)
                {
                    ShowErrorImage($"Export failed: {ex.Message}");
                    return;
                }
                
                if (exported)
                {
                    string expectedFile = filePath + ".png";
                    
                    // Handle Revit's naming convention
                    if (!File.Exists(expectedFile))
                    {
                        string[] possibleFiles = new string[]
                        {
                            filePath + " - " + view.Name + ".png",
                            filePath + "_" + view.Name + ".png",
                            filePath + " - " + view.ViewType + " - " + view.Name + ".png"
                        };
                        
                        foreach (string possibleFile in possibleFiles)
                        {
                            if (File.Exists(possibleFile))
                            {
                                expectedFile = possibleFile;
                                break;
                            }
                        }
                        
                        // If still not found, search for any file with the base name
                        if (!File.Exists(expectedFile))
                        {
                            string[] files = Directory.GetFiles(tempPath, fileName + "*.png");
                            if (files.Length > 0)
                            {
                                expectedFile = files[0];
                            }
                        }
                    }
                    
                    if (File.Exists(expectedFile))
                    {
                        // Load image
                        using (var stream = new FileStream(expectedFile, FileMode.Open, FileAccess.Read))
                        {
                            var img = Image.FromStream(stream);
                            originalImage = new Bitmap(img);
                            pictureBox.Image = originalImage;
                        }
                        
                        statusLabel.Text = "Preview loaded - Scroll to zoom, double-click to fit";
                        
                        // Clean up temp file
                        try { File.Delete(expectedFile); } catch { }
                    }
                    else
                    {
                        ShowErrorImage("Preview file not found");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowErrorImage($"Error: {ex.Message}");
            }
        }
        
        private void ShowScheduleMessage()
        {
            statusLabel.Text = "Schedule views cannot be previewed";
            
            try
            {
                using (Bitmap scheduleImg = new Bitmap(400, 300))
                using (Graphics g = Graphics.FromImage(scheduleImg))
                {
                    g.Clear(System.Drawing.Color.FromArgb(250, 250, 250));
                    
                    // Draw schedule icon
                    int centerX = 200;
                    int centerY = 120;
                    
                    using (Pen pen = new Pen(System.Drawing.Color.FromArgb(100, 100, 100), 2))
                    {
                        // Draw table grid
                        for (int i = 0; i < 4; i++)
                        {
                            int y = centerY - 40 + (i * 20);
                            g.DrawLine(pen, centerX - 60, y, centerX + 60, y);
                        }
                        for (int i = 0; i < 4; i++)
                        {
                            int x = centerX - 60 + (i * 40);
                            g.DrawLine(pen, x, centerY - 40, x, centerY + 20);
                        }
                    }
                    
                    using (Font font = new Font("Segoe UI", 12, FontStyle.Bold))
                    using (Brush brush = new SolidBrush(System.Drawing.Color.FromArgb(80, 80, 80)))
                    {
                        string text = "Schedule View";
                        SizeF textSize = g.MeasureString(text, font);
                        float x = (400 - textSize.Width) / 2;
                        float y = centerY + 40;
                        g.DrawString(text, font, brush, x, y);
                        
                        using (Font smallFont = new Font("Segoe UI", 10))
                        {
                            string subText = "Preview not available for schedules";
                            textSize = g.MeasureString(subText, smallFont);
                            x = (400 - textSize.Width) / 2;
                            y += 25;
                            g.DrawString(subText, smallFont, brush, x, y);
                        }
                    }
                    
                    pictureBox.Image = new Bitmap(scheduleImg);
                }
            }
            catch { }
        }
        
        private void ShowErrorImage(string errorMessage)
        {
            statusLabel.Text = errorMessage;
            
            try
            {
                using (Bitmap errorImg = new Bitmap(400, 300))
                using (Graphics g = Graphics.FromImage(errorImg))
                {
                    g.Clear(System.Drawing.Color.FromArgb(250, 250, 250));
                    
                    // Draw error icon
                    int centerX = 200;
                    int centerY = 120;
                    int size = 60;
                    
                    using (Pen pen = new Pen(System.Drawing.Color.FromArgb(220, 50, 47), 4))
                    {
                        g.DrawEllipse(pen, centerX - size/2, centerY - size/2, size, size);
                        g.DrawLine(pen, centerX - 15, centerY - 15, centerX + 15, centerY + 15);
                        g.DrawLine(pen, centerX - 15, centerY + 15, centerX + 15, centerY - 15);
                    }
                    
                    using (Font font = new Font("Segoe UI", 11))
                    using (Brush brush = new SolidBrush(System.Drawing.Color.FromArgb(100, 100, 100)))
                    {
                        string text = "Preview not available";
                        SizeF textSize = g.MeasureString(text, font);
                        float x = (400 - textSize.Width) / 2;
                        float y = centerY + size/2 + 20;
                        g.DrawString(text, font, brush, x, y);
                        
                        using (Font smallFont = new Font("Segoe UI", 9))
                        {
                            textSize = g.MeasureString(errorMessage, smallFont);
                            x = (400 - textSize.Width) / 2;
                            y += 25;
                            g.DrawString(errorMessage, smallFont, brush, x, y);
                        }
                    }
                    
                    pictureBox.Image = new Bitmap(errorImg);
                }
            }
            catch { }
        }
        
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            // Don't call StopMonitoring here - it causes circular reference
            // The form is already being closed
            
            if (pictureBox.Image != null)
            {
                pictureBox.Image.Dispose();
                pictureBox.Image = null;
            }
            
            if (originalImage != null)
            {
                originalImage.Dispose();
                originalImage = null;
            }
            
            base.OnFormClosed(e);
        }
    }
}