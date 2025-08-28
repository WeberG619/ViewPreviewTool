using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System.Windows.Media.Imaging;

namespace ViewPreviewTool
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class ViewPreviewApplication : IExternalApplication
    {
        private static ViewPreviewForm _previewWindow = null;
        private static UIApplication _uiApp = null;
        private static bool _isEnabled = false;
        private static System.Windows.Forms.Timer _hoverTimer = null;
        private static System.Windows.Forms.Timer _closeTimer = null;
        private static ElementId _lastHoveredId = ElementId.InvalidElementId;
        private static DateTime _lastMouseMove = DateTime.Now;
        private static Point _lastMousePosition = Point.Empty;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create ribbon panel
                RibbonPanel ribbonPanel = application.CreateRibbonPanel("BIM Ops Studio");
                
                // Create button
                PushButtonData buttonData = new PushButtonData(
                    "ViewPreviewTool",
                    "View\nPreview",
                    System.Reflection.Assembly.GetExecutingAssembly().Location,
                    "ViewPreviewTool.ViewPreviewCommand");
                
                buttonData.LargeImage = CreateButtonImage(32);
                buttonData.Image = CreateButtonImage(16);
                buttonData.ToolTip = "Toggle View Preview Tool";
                buttonData.LongDescription = "Enable/disable hover preview for views in the Project Browser.";
                
                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;
                
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            Cleanup();
            return Result.Succeeded;
        }

        private static System.Windows.Media.Imaging.BitmapSource CreateButtonImage(int size)
        {
            try
            {
                using (Bitmap bitmap = new Bitmap(size, size))
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.Transparent);
                    
                    using (Brush brush = new LinearGradientBrush(
                        new Rectangle(0, 0, size, size),
                        Color.FromArgb(0, 122, 204),
                        Color.FromArgb(0, 88, 156),
                        45F))
                    {
                        g.FillEllipse(brush, 2, 2, size - 4, size - 4);
                    }
                    
                    using (Pen pen = new Pen(Color.White, size / 8))
                    {
                        int margin = size / 4;
                        g.DrawRectangle(pen, margin, margin, size - 2 * margin, size - 2 * margin);
                    }
                    
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                        bitmap.GetHbitmap(),
                        IntPtr.Zero,
                        System.Windows.Int32Rect.Empty,
                        System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());
                }
            }
            catch
            {
                return null;
            }
        }

        [Transaction(TransactionMode.ReadOnly)]
        public class ViewPreviewCommand : IExternalCommand
        {
            public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
            {
                try
                {
                    _uiApp = commandData.Application;
                    
                    if (_isEnabled)
                    {
                        // Disable
                        _isEnabled = false;
                        _uiApp.Idling -= OnIdling;
                        Cleanup();
                        TaskDialog.Show("View Preview", "View Preview Tool disabled.");
                    }
                    else
                    {
                        // Enable
                        _isEnabled = true;
                        _uiApp.Idling += OnIdling;
                        TaskDialog.Show("View Preview", "View Preview Tool enabled.\n\nHover over views in the Project Browser to see previews.");
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

        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            if (!_isEnabled || _uiApp == null) return;
            
            try
            {
                UIDocument uidoc = _uiApp.ActiveUIDocument;
                if (uidoc == null) return;
                
                // Check if mouse has moved significantly
                Point currentMousePos = Cursor.Position;
                bool mouseMoved = Math.Abs(currentMousePos.X - _lastMousePosition.X) > 5 || 
                                  Math.Abs(currentMousePos.Y - _lastMousePosition.Y) > 5;
                
                if (mouseMoved)
                {
                    _lastMousePosition = currentMousePos;
                    _lastMouseMove = DateTime.Now;
                }
                
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 1)
                {
                    ElementId currentId = selectedIds.First();
                    Element elem = uidoc.Document.GetElement(currentId);
                    
                    if (elem is View view && !(elem is ViewSheet))
                    {
                        if (!currentId.Equals(_lastHoveredId))
                        {
                            _lastHoveredId = currentId;
                            StartHoverTimer(view, uidoc.Document);
                        }
                        
                        // Keep window open if hovering over same view
                        if (_previewWindow != null && !_previewWindow.IsDisposed)
                        {
                            ResetCloseTimer();
                        }
                    }
                    else
                    {
                        // Not a view, start close timer
                        if (_previewWindow != null && !_previewWindow.IsDisposed)
                        {
                            StartCloseTimer();
                        }
                        _lastHoveredId = ElementId.InvalidElementId;
                    }
                }
                else
                {
                    // No selection or multiple selections
                    if (_previewWindow != null && !_previewWindow.IsDisposed)
                    {
                        StartCloseTimer();
                    }
                    _lastHoveredId = ElementId.InvalidElementId;
                }
            }
            catch
            {
                // Silently handle errors
            }
        }

        private static void StartHoverTimer(View view, Document doc)
        {
            CleanupHoverTimer();
            
            _hoverTimer = new System.Windows.Forms.Timer();
            _hoverTimer.Interval = 800; // 0.8 second delay before showing
            _hoverTimer.Tick += (s, e) =>
            {
                CleanupHoverTimer();
                ShowPreview(view, doc);
            };
            _hoverTimer.Start();
        }

        private static void StartCloseTimer()
        {
            if (_closeTimer == null)
            {
                _closeTimer = new System.Windows.Forms.Timer();
                _closeTimer.Interval = 500; // 0.5 second delay before closing
                _closeTimer.Tick += (s, e) =>
                {
                    CleanupCloseTimer();
                    ClosePreviewWindow();
                };
                _closeTimer.Start();
            }
        }

        private static void ResetCloseTimer()
        {
            CleanupCloseTimer();
        }

        private static void CleanupHoverTimer()
        {
            if (_hoverTimer != null)
            {
                _hoverTimer.Stop();
                _hoverTimer.Dispose();
                _hoverTimer = null;
            }
        }

        private static void CleanupCloseTimer()
        {
            if (_closeTimer != null)
            {
                _closeTimer.Stop();
                _closeTimer.Dispose();
                _closeTimer = null;
            }
        }

        private static void Cleanup()
        {
            CleanupHoverTimer();
            CleanupCloseTimer();
            ClosePreviewWindow();
        }

        private static void ShowPreview(View view, Document doc)
        {
            try
            {
                // Don't create new window if one already exists
                if (_previewWindow != null && !_previewWindow.IsDisposed)
                {
                    _previewWindow.UpdateView(view, doc);
                    return;
                }
                
                _previewWindow = new ViewPreviewForm(view, doc);
                _previewWindow.TopMost = true;
                _previewWindow.ShowInTaskbar = false;
                
                // Position near cursor but not under it
                Point cursorPos = Cursor.Position;
                _previewWindow.StartPosition = FormStartPosition.Manual;
                
                // Ensure window stays on screen
                Screen screen = Screen.FromPoint(cursorPos);
                int x = cursorPos.X + 20;
                int y = cursorPos.Y - _previewWindow.Height / 2;
                
                // Adjust if window would go off screen
                if (x + _previewWindow.Width > screen.WorkingArea.Right)
                    x = cursorPos.X - _previewWindow.Width - 20;
                    
                if (y < screen.WorkingArea.Top)
                    y = screen.WorkingArea.Top;
                else if (y + _previewWindow.Height > screen.WorkingArea.Bottom)
                    y = screen.WorkingArea.Bottom - _previewWindow.Height;
                
                _previewWindow.Location = new Point(x, y);
                _previewWindow.Show();
                
                // Track mouse to keep window open while hovering
                _previewWindow.MouseEnter += (s, e) => ResetCloseTimer();
                _previewWindow.MouseLeave += (s, e) => StartCloseTimer();
            }
            catch
            {
                ClosePreviewWindow();
            }
        }

        private static void ClosePreviewWindow()
        {
            if (_previewWindow != null && !_previewWindow.IsDisposed)
            {
                _previewWindow.Close();
                _previewWindow.Dispose();
            }
            _previewWindow = null;
        }

        public class ViewPreviewForm : Form
        {
            private PictureBox pictureBox;
            private Label titleLabel;
            private Label typeLabel;
            private Panel headerPanel;
            private View currentView;
            private Document currentDoc;
            
            public ViewPreviewForm(View view, Document doc)
            {
                InitializeForm();
                UpdateView(view, doc);
            }
            
            public void UpdateView(View view, Document doc)
            {
                currentView = view;
                currentDoc = doc;
                titleLabel.Text = view.Name;
                typeLabel.Text = $"Type: {view.ViewType}";
                LoadViewImage();
            }
            
            private void InitializeForm()
            {
                this.Text = "View Preview";
                this.Size = new Size(700, 600);
                this.FormBorderStyle = FormBorderStyle.None;
                this.StartPosition = FormStartPosition.Manual;
                this.BackColor = Color.FromArgb(45, 45, 48);
                this.ShowIcon = false;
                
                // Add border
                this.Paint += (s, e) =>
                {
                    using (Pen pen = new Pen(Color.FromArgb(0, 122, 204), 2))
                    {
                        e.Graphics.DrawRectangle(pen, 0, 0, this.Width - 1, this.Height - 1);
                    }
                };
                
                // Header panel
                headerPanel = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 80,
                    BackColor = Color.FromArgb(0, 122, 204)
                };
                
                // Title label
                titleLabel = new Label
                {
                    Location = new Point(15, 15),
                    Size = new Size(670, 30),
                    Font = new Font("Segoe UI", 16, FontStyle.Bold),
                    ForeColor = Color.White,
                    BackColor = Color.Transparent,
                    Text = "View Name"
                };
                headerPanel.Controls.Add(titleLabel);
                
                // Type label
                typeLabel = new Label
                {
                    Location = new Point(15, 45),
                    Size = new Size(670, 25),
                    Font = new Font("Segoe UI", 11),
                    ForeColor = Color.FromArgb(220, 220, 220),
                    BackColor = Color.Transparent,
                    Text = "View Type"
                };
                headerPanel.Controls.Add(typeLabel);
                
                // Close button
                Label closeButton = new Label
                {
                    Text = "✕",
                    Font = new Font("Arial", 14, FontStyle.Bold),
                    ForeColor = Color.White,
                    BackColor = Color.Transparent,
                    Size = new Size(30, 30),
                    Location = new Point(this.Width - 40, 10),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand
                };
                closeButton.Click += (s, e) => this.Close();
                closeButton.MouseEnter += (s, e) => closeButton.BackColor = Color.FromArgb(229, 20, 0);
                closeButton.MouseLeave += (s, e) => closeButton.BackColor = Color.Transparent;
                headerPanel.Controls.Add(closeButton);
                
                // Picture box container
                Panel containerPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = Color.White,
                    Padding = new Padding(2)
                };
                
                // Picture box
                pictureBox = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = Color.White
                };
                containerPanel.Controls.Add(pictureBox);
                
                // Add controls
                this.Controls.Add(containerPanel);
                this.Controls.Add(headerPanel);
                
                // Enable dragging
                bool dragging = false;
                Point dragCursorPoint = Point.Empty;
                Point dragFormPoint = Point.Empty;
                
                headerPanel.MouseDown += (s, e) =>
                {
                    dragging = true;
                    dragCursorPoint = Cursor.Position;
                    dragFormPoint = this.Location;
                };
                
                headerPanel.MouseMove += (s, e) =>
                {
                    if (dragging)
                    {
                        Point diff = Point.Subtract(Cursor.Position, new Size(dragCursorPoint));
                        this.Location = Point.Add(dragFormPoint, new Size(diff));
                    }
                };
                
                headerPanel.MouseUp += (s, e) => dragging = false;
            }
            
            private void LoadViewImage()
            {
                try
                {
                    // Show loading text
                    using (Bitmap loadingImg = new Bitmap(pictureBox.Width, pictureBox.Height))
                    using (Graphics g = Graphics.FromImage(loadingImg))
                    {
                        g.Clear(Color.White);
                        using (Font font = new Font("Segoe UI", 14))
                        {
                            g.DrawString("Loading preview...", font, Brushes.Gray, 
                                new RectangleF(0, 0, loadingImg.Width, loadingImg.Height), 
                                new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                        }
                        pictureBox.Image = (Bitmap)loadingImg.Clone();
                    }
                    
                    Application.DoEvents();
                    
                    string tempPath = Path.GetTempPath();
                    string fileName = $"ViewPreview_{currentView.Id.Value}_{DateTime.Now.Ticks}";
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
                    
                    options.SetViewsAndSheets(new List<ElementId> { currentView.Id });
                    
                    if (currentDoc.Export(tempPath, fileName, options))
                    {
                        string expectedFile = filePath + ".png";
                        if (!File.Exists(expectedFile))
                        {
                            string[] files = Directory.GetFiles(tempPath, fileName + "*.png");
                            if (files.Length > 0)
                                expectedFile = files[0];
                        }
                        
                        if (File.Exists(expectedFile))
                        {
                            using (var stream = new FileStream(expectedFile, FileMode.Open, FileAccess.Read))
                            {
                                if (pictureBox.Image != null)
                                {
                                    pictureBox.Image.Dispose();
                                }
                                pictureBox.Image = Image.FromStream(stream);
                            }
                            
                            try { File.Delete(expectedFile); } catch { }
                        }
                    }
                    else
                    {
                        ShowErrorImage("Unable to export view");
                    }
                }
                catch (Exception ex)
                {
                    ShowErrorImage($"Error: {ex.Message}");
                }
            }
            
            private void ShowErrorImage(string message)
            {
                using (Bitmap errorImg = new Bitmap(600, 400))
                using (Graphics g = Graphics.FromImage(errorImg))
                {
                    g.Clear(Color.FromArgb(245, 245, 245));
                    using (Font font = new Font("Segoe UI", 12))
                    {
                        g.DrawString(message, font, Brushes.DarkGray, 
                            new RectangleF(0, 0, 600, 400), 
                            new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                    }
                    if (pictureBox.Image != null)
                    {
                        pictureBox.Image.Dispose();
                    }
                    pictureBox.Image = (Bitmap)errorImg.Clone();
                }
            }
            
            protected override void OnFormClosed(FormClosedEventArgs e)
            {
                if (pictureBox.Image != null)
                {
                    pictureBox.Image.Dispose();
                    pictureBox.Image = null;
                }
                base.OnFormClosed(e);
            }
            
            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams cp = base.CreateParams;
                    cp.ClassStyle |= 0x00020000; // CS_DROPSHADOW
                    return cp;
                }
            }
        }
    }
}