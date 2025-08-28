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

namespace ViewPreviewTool
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class ViewPreviewApplication : IExternalApplication
    {
        private static Form _previewWindow = null;
        private static UIApplication _uiApp = null;
        private static bool _isHovering = false;
        private static System.Windows.Forms.Timer _hoverTimer = null;
        private static ElementId _lastHoveredId = ElementId.InvalidElementId;

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
            CleanupTimer();
            ClosePreviewWindow();
            return Result.Succeeded;
        }

        private static BitmapSource CreateButtonImage(int size)
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
                        BitmapSizeOptions.FromEmptyOptions());
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
                    
                    if (_isHovering)
                    {
                        // Disable
                        _isHovering = false;
                        _uiApp.Idling -= OnIdling;
                        CleanupTimer();
                        ClosePreviewWindow();
                        TaskDialog.Show("View Preview", "View Preview Tool disabled.");
                    }
                    else
                    {
                        // Enable
                        _isHovering = true;
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
            if (!_isHovering || _uiApp == null) return;
            
            try
            {
                UIDocument uidoc = _uiApp.ActiveUIDocument;
                if (uidoc == null) return;
                
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 1)
                {
                    ElementId currentId = selectedIds.First();
                    if (!currentId.Equals(_lastHoveredId))
                    {
                        _lastHoveredId = currentId;
                        Element elem = uidoc.Document.GetElement(currentId);
                        
                        if (elem is View view && !(elem is ViewSheet))
                        {
                            StartHoverTimer(view, uidoc.Document);
                        }
                        else
                        {
                            CleanupTimer();
                            ClosePreviewWindow();
                        }
                    }
                }
                else
                {
                    _lastHoveredId = ElementId.InvalidElementId;
                    CleanupTimer();
                    ClosePreviewWindow();
                }
            }
            catch
            {
                // Silently handle errors
            }
        }

        private static void StartHoverTimer(View view, Document doc)
        {
            CleanupTimer();
            
            _hoverTimer = new System.Windows.Forms.Timer();
            _hoverTimer.Interval = 500; // 0.5 second delay
            _hoverTimer.Tick += (s, e) =>
            {
                CleanupTimer();
                ShowPreview(view, doc);
            };
            _hoverTimer.Start();
        }

        private static void CleanupTimer()
        {
            if (_hoverTimer != null)
            {
                _hoverTimer.Stop();
                _hoverTimer.Dispose();
                _hoverTimer = null;
            }
        }

        private static void ShowPreview(View view, Document doc)
        {
            try
            {
                ClosePreviewWindow();
                
                _previewWindow = new ViewPreviewForm(view, doc);
                _previewWindow.TopMost = false;
                _previewWindow.ShowInTaskbar = false;
                
                // Position near cursor
                Point cursorPos = Cursor.Position;
                _previewWindow.StartPosition = FormStartPosition.Manual;
                
                // Ensure window stays on screen
                Screen screen = Screen.FromPoint(cursorPos);
                int x = Math.Min(cursorPos.X + 20, screen.WorkingArea.Right - _previewWindow.Width);
                int y = Math.Min(cursorPos.Y + 20, screen.WorkingArea.Bottom - _previewWindow.Height);
                x = Math.Max(x, screen.WorkingArea.Left);
                y = Math.Max(y, screen.WorkingArea.Top);
                
                _previewWindow.Location = new Point(x, y);
                _previewWindow.Show();
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
            
            public ViewPreviewForm(View view, Document doc)
            {
                InitializeForm();
                LoadViewImage(view, doc);
            }
            
            private void InitializeForm()
            {
                this.Text = "View Preview";
                this.Size = new Size(600, 500);
                this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
                this.StartPosition = FormStartPosition.Manual;
                this.BackColor = Color.White;
                
                pictureBox = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = Color.White
                };
                
                this.Controls.Add(pictureBox);
            }
            
            private void LoadViewImage(View view, Document doc)
            {
                try
                {
                    string tempPath = Path.GetTempPath();
                    string fileName = $"ViewPreview_{view.Id.Value}_{DateTime.Now.Ticks}";
                    string filePath = Path.Combine(tempPath, fileName);
                    
                    var options = new ImageExportOptions
                    {
                        FilePath = filePath,
                        FitDirection = FitDirectionType.Horizontal,
                        HLRandWFViewsFileType = ImageFileType.PNG,
                        ImageResolution = ImageResolution.DPI_150,
                        PixelSize = 800,
                        ExportRange = ExportRange.SetOfViews
                    };
                    
                    options.SetViewsAndSheets(new List<ElementId> { view.Id });
                    
                    if (doc.Export(tempPath, fileName, options))
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
                                pictureBox.Image = Image.FromStream(stream);
                            }
                            
                            try { File.Delete(expectedFile); } catch { }
                        }
                    }
                }
                catch
                {
                    // Show error image
                    using (Bitmap errorImg = new Bitmap(400, 300))
                    using (Graphics g = Graphics.FromImage(errorImg))
                    {
                        g.Clear(Color.LightGray);
                        using (Font font = new Font("Arial", 12))
                        {
                            g.DrawString("Preview not available", font, Brushes.DarkGray, 
                                new RectangleF(0, 0, 400, 300), 
                                new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                        }
                    }
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
        }
    }
    
    // Add missing using directive
    using BitmapSource = System.Windows.Media.Imaging.BitmapSource;
    using BitmapSizeOptions = System.Windows.Media.Imaging.BitmapSizeOptions;
}