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

namespace ViewPreviewTool
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class ViewPreviewApplication : IExternalApplication
    {
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
                    "ViewPreviewTool.ShowPreviewCommand");
                
                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;
                pushButton.ToolTip = "Show preview of selected view";
                pushButton.LongDescription = "Select a view in the Project Browser and click to see a preview.";
                
                // Try to set icon
                try
                {
                    pushButton.LargeImage = CreateButtonImage(32);
                    pushButton.Image = CreateButtonImage(16);
                }
                catch { }
                
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
        
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        
        private static System.Windows.Media.Imaging.BitmapSource CreateButtonImage(int size)
        {
            using (Bitmap bitmap = new Bitmap(size, size))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(System.Drawing.Color.Transparent);
                
                using (Brush brush = new LinearGradientBrush(
                    new System.Drawing.Rectangle(0, 0, size, size),
                    System.Drawing.Color.FromArgb(0, 122, 204),
                    System.Drawing.Color.FromArgb(0, 88, 156),
                    45F))
                {
                    g.FillEllipse(brush, 2, 2, size - 4, size - 4);
                }
                
                using (System.Drawing.Pen pen = new System.Drawing.Pen(System.Drawing.Color.White, size / 8))
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
    }
    
    [Transaction(TransactionMode.ReadOnly)]
    public class ShowPreviewCommand : IExternalCommand
    {
        private static ViewPreviewForm previewForm = null;
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                if (uidoc == null)
                {
                    message = "No active document";
                    return Result.Failed;
                }
                
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 0)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("View Preview", 
                        "Please select a view in the Project Browser first, then click View Preview.");
                    return Result.Cancelled;
                }
                
                if (selectedIds.Count > 1)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("View Preview", 
                        "Please select only one view to preview.");
                    return Result.Cancelled;
                }
                
                ElementId selectedId = selectedIds.First();
                Element elem = uidoc.Document.GetElement(selectedId);
                
                if (elem == null)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("View Preview", 
                        "Could not find the selected element.");
                    return Result.Failed;
                }
                
                Autodesk.Revit.DB.View view = elem as Autodesk.Revit.DB.View;
                
                if (view == null)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("View Preview", 
                        "Selected element is not a view.\n\nPlease select a view from the Project Browser.");
                    return Result.Failed;
                }
                
                // Check if it's a valid view type
                if (view is ViewSheet)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("View Preview", 
                        "Sheet views cannot be previewed.\n\nPlease select a model view.");
                    return Result.Failed;
                }
                
                // Check if view can be exported
                if (!view.CanBePrinted)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("View Preview", 
                        "This view cannot be exported for preview.\n\nTry selecting a different view.");
                    return Result.Failed;
                }
                
                // Close existing preview if open
                if (previewForm != null && !previewForm.IsDisposed)
                {
                    previewForm.Close();
                    previewForm.Dispose();
                    previewForm = null;
                }
                
                // Create and show new preview
                previewForm = new ViewPreviewForm(uidoc.Document, view);
                previewForm.Show();
                
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
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Timer loadTimer;
        
        public ViewPreviewForm(Document document, Autodesk.Revit.DB.View v)
        {
            doc = document;
            view = v;
            InitializeForm();
            
            // Use timer to load image after form is shown
            loadTimer = new System.Windows.Forms.Timer();
            loadTimer.Interval = 100;
            loadTimer.Tick += (s, e) =>
            {
                loadTimer.Stop();
                loadTimer.Dispose();
                LoadViewImage();
            };
            loadTimer.Start();
        }
        
        private void InitializeForm()
        {
            this.Text = $"View Preview: {view.Name}";
            this.Size = new Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(600, 500);
            this.ShowIcon = false;
            
            // Header panel
            System.Windows.Forms.Panel headerPanel = new System.Windows.Forms.Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 80;
            headerPanel.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            
            // Title label
            System.Windows.Forms.Label titleLabel = new System.Windows.Forms.Label();
            titleLabel.Text = TruncateText(view.Name, 50);
            titleLabel.Font = new Font("Segoe UI", 18, FontStyle.Bold);
            titleLabel.ForeColor = System.Drawing.Color.White;
            titleLabel.Location = new System.Drawing.Point(20, 15);
            titleLabel.AutoSize = false;
            titleLabel.Size = new Size(860, 35);
            titleLabel.BackColor = System.Drawing.Color.Transparent;
            headerPanel.Controls.Add(titleLabel);
            
            // Type label
            System.Windows.Forms.Label typeLabel = new System.Windows.Forms.Label();
            typeLabel.Text = $"Type: {view.ViewType}";
            typeLabel.Font = new Font("Segoe UI", 11);
            typeLabel.ForeColor = System.Drawing.Color.FromArgb(220, 220, 220);
            typeLabel.Location = new System.Drawing.Point(20, 50);
            typeLabel.AutoSize = false;
            typeLabel.Size = new Size(860, 25);
            typeLabel.BackColor = System.Drawing.Color.Transparent;
            headerPanel.Controls.Add(typeLabel);
            
            // Status panel
            System.Windows.Forms.Panel statusPanel = new System.Windows.Forms.Panel();
            statusPanel.Dock = DockStyle.Bottom;
            statusPanel.Height = 30;
            statusPanel.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            
            statusLabel = new System.Windows.Forms.Label();
            statusLabel.Text = "Loading preview...";
            statusLabel.Dock = DockStyle.Fill;
            statusLabel.TextAlign = ContentAlignment.MiddleCenter;
            statusLabel.Font = new Font("Segoe UI", 9);
            statusLabel.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80);
            statusPanel.Controls.Add(statusLabel);
            
            // Picture box container
            System.Windows.Forms.Panel containerPanel = new System.Windows.Forms.Panel();
            containerPanel.Dock = DockStyle.Fill;
            containerPanel.BackColor = System.Drawing.Color.White;
            containerPanel.Padding = new Padding(10);
            
            // Picture box
            pictureBox = new PictureBox();
            pictureBox.Dock = DockStyle.Fill;
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox.BackColor = System.Drawing.Color.White;
            
            // Show loading image
            ShowLoadingImage();
            
            containerPanel.Controls.Add(pictureBox);
            
            // Add controls
            this.Controls.Add(containerPanel);
            this.Controls.Add(statusPanel);
            this.Controls.Add(headerPanel);
            
            // Add zoom functionality on double-click
            pictureBox.DoubleClick += (s, e) =>
            {
                if (pictureBox.Image == null) return;
                
                if (pictureBox.SizeMode == PictureBoxSizeMode.Zoom)
                {
                    pictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
                    containerPanel.AutoScroll = true;
                }
                else
                {
                    pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
                    containerPanel.AutoScroll = false;
                }
            };
        }
        
        private string TruncateText(string text, int maxLength)
        {
            if (text.Length <= maxLength) return text;
            return text.Substring(0, maxLength - 3) + "...";
        }
        
        private void ShowLoadingImage()
        {
            try
            {
                using (Bitmap loadingImg = new Bitmap(600, 400))
                using (Graphics g = Graphics.FromImage(loadingImg))
                {
                    g.Clear(System.Drawing.Color.White);
                    using (Font font = new Font("Segoe UI", 14))
                    {
                        string text = "Generating preview...";
                        SizeF textSize = g.MeasureString(text, font);
                        g.DrawString(text, font, Brushes.Gray, 
                            (600 - textSize.Width) / 2, 
                            (400 - textSize.Height) / 2);
                    }
                    pictureBox.Image = (Bitmap)loadingImg.Clone();
                }
            }
            catch { }
        }
        
        private void LoadViewImage()
        {
            try
            {
                statusLabel.Text = "Generating preview...";
                Application.DoEvents();
                
                string tempPath = Path.GetTempPath();
                string fileName = $"ViewPreview_{view.Id.Value}_{DateTime.Now.Ticks}";
                string filePath = Path.Combine(tempPath, fileName);
                
                var options = new ImageExportOptions();
                options.FilePath = filePath;
                options.FitDirection = FitDirectionType.Horizontal;
                options.HLRandWFViewsFileType = ImageFileType.PNG;
                options.ImageResolution = ImageResolution.DPI_300;
                options.PixelSize = 1600;
                options.ExportRange = ExportRange.SetOfViews;
                options.SetViewsAndSheets(new List<ElementId> { view.Id });
                
                bool exported = false;
                try
                {
                    // In Revit 2025, use the Export method without casting
                    exported = doc.Export(tempPath, fileName, options);
                }
                catch (Exception ex)
                {
                    ShowErrorImage($"Export failed: {ex.Message}");
                    return;
                }
                
                if (exported)
                {
                    string expectedFile = filePath + ".png";
                    
                    // Sometimes Revit adds numbers to the filename
                    if (!File.Exists(expectedFile))
                    {
                        string[] files = Directory.GetFiles(tempPath, fileName + "*.png");
                        if (files.Length > 0)
                        {
                            expectedFile = files[0];
                        }
                    }
                    
                    if (File.Exists(expectedFile))
                    {
                        // Load the image
                        using (var stream = new FileStream(expectedFile, FileMode.Open, FileAccess.Read))
                        {
                            var img = Image.FromStream(stream);
                            if (pictureBox.Image != null)
                            {
                                pictureBox.Image.Dispose();
                            }
                            pictureBox.Image = img;
                        }
                        
                        statusLabel.Text = "Preview loaded. Double-click image to toggle zoom mode.";
                        
                        // Clean up temp file
                        try 
                        { 
                            System.Threading.Thread.Sleep(100);
                            File.Delete(expectedFile); 
                        } 
                        catch { }
                    }
                    else
                    {
                        ShowErrorImage("Preview file not found after export");
                    }
                }
                else
                {
                    ShowErrorImage("Unable to export view - export returned false");
                }
            }
            catch (Exception ex)
            {
                ShowErrorImage($"Error: {ex.Message}");
            }
        }
        
        private void ShowErrorImage(string errorMessage)
        {
            statusLabel.Text = errorMessage;
            
            try
            {
                using (Bitmap errorImg = new Bitmap(600, 400))
                using (Graphics g = Graphics.FromImage(errorImg))
                {
                    g.Clear(System.Drawing.Color.FromArgb(245, 245, 245));
                    
                    using (Font titleFont = new Font("Segoe UI", 16))
                    using (Font msgFont = new Font("Segoe UI", 11))
                    {
                        string title = "Preview Not Available";
                        SizeF titleSize = g.MeasureString(title, titleFont);
                        g.DrawString(title, titleFont, Brushes.DarkGray, 
                            (600 - titleSize.Width) / 2, 150);
                        
                        // Word wrap the error message
                        System.Drawing.Rectangle msgRect = new System.Drawing.Rectangle(50, 200, 500, 150);
                        g.DrawString(errorMessage, msgFont, Brushes.Gray, msgRect);
                    }
                    
                    if (pictureBox.Image != null)
                    {
                        pictureBox.Image.Dispose();
                    }
                    pictureBox.Image = (Bitmap)errorImg.Clone();
                }
            }
            catch { }
        }
        
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (loadTimer != null)
            {
                loadTimer.Stop();
                loadTimer.Dispose();
            }
            
            if (pictureBox.Image != null)
            {
                pictureBox.Image.Dispose();
                pictureBox.Image = null;
            }
            base.OnFormClosed(e);
        }
    }
}