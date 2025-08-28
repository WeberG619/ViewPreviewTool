using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Resolve ambiguities
using Form = System.Windows.Forms.Form;
using View = Autodesk.Revit.DB.View;
using Control = System.Windows.Forms.Control;
using Point = System.Drawing.Point;
using Panel = System.Windows.Forms.Panel;
using Color = System.Drawing.Color;

namespace ViewPreviewTool
{
    [Transaction(TransactionMode.Manual)]
    public class ViewPreviewCommand : IExternalCommand
    {
        private static Form previewForm = null;
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                if (uidoc == null) return Result.Failed;
                
                var selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count != 1)
                {
                    TaskDialog.Show("View Preview", "Please select a single view to preview.");
                    return Result.Cancelled;
                }
                
                Element elem = null;
                foreach (ElementId id in selectedIds)
                {
                    elem = uidoc.Document.GetElement(id);
                    break;
                }
                
                View view = elem as View;
                
                if (view == null)
                {
                    TaskDialog.Show("View Preview", "Selected element is not a view.");
                    return Result.Failed;
                }
                
                // Close existing preview
                if (previewForm != null && !previewForm.IsDisposed)
                {
                    previewForm.Close();
                    previewForm.Dispose();
                }
                
                // Create new preview with TopMost = false
                previewForm = new ViewPreviewForm(uidoc.Document, view);
                previewForm.TopMost = false;  // Fix the always-on-top issue
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
    
    public class ViewPreviewForm : Form
    {
        private Document doc;
        private View view;
        private ZoomablePictureBox pictureBox;
        private Label titleLabel;
        
        public ViewPreviewForm(Document document, View v)
        {
            doc = document;
            view = v;
            InitializeForm();
            LoadViewImage();
        }
        
        private void InitializeForm()
        {
            this.Text = string.Format("View Preview: {0} - v2.5", view.Name);
            this.Size = new Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = false;  // Ensure window doesn't stay on top
            
            // Try to load BIM Ops Studio icon
            try
            {
                string logoPath = @"D:\BIM_Ops_Studio\BIM_Ops_Studio_logo_all_sizes\BIM_Ops_Studio_logo_1500x1500.png";
                if (File.Exists(logoPath))
                {
                    using (var img = Image.FromFile(logoPath))
                    {
                        var resized = new Bitmap(img, new Size(32, 32));
                        this.Icon = Icon.FromHandle(resized.GetHicon());
                    }
                }
            }
            catch { }
            
            // Header panel - BIGGER and BETTER
            Panel headerPanel = new Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 80;  // Increased from 60 to 80
            headerPanel.Paint += HeaderPanel_Paint;
            
            // Logo
            try
            {
                string logoPath = @"D:\BIM_Ops_Studio\BIM_Ops_Studio_logo_all_sizes\BIM_Ops_Studio_logo_1500x1500.png";
                if (File.Exists(logoPath))
                {
                    PictureBox logoPic = new PictureBox();
                    using (var img = Image.FromFile(logoPath))
                    {
                        var logo = new Bitmap(60, 60);  // Bigger logo
                        using (var g = Graphics.FromImage(logo))
                        {
                            g.SmoothingMode = SmoothingMode.AntiAlias;
                            g.Clear(Color.Transparent);
                            g.DrawImage(img, 0, 0, 60, 60);
                        }
                        logoPic.Image = logo;
                    }
                    logoPic.SizeMode = PictureBoxSizeMode.Zoom;
                    logoPic.Size = new Size(60, 60);
                    logoPic.Location = new Point(15, 10);
                    logoPic.BackColor = Color.Transparent;
                    headerPanel.Controls.Add(logoPic);
                }
            }
            catch { }
            
            // Title - BIGGER and WHITER
            titleLabel = new Label();
            titleLabel.Text = TruncateText(view.Name, 50);
            titleLabel.Font = new Font("Segoe UI", 18, FontStyle.Bold);  // Increased from 14 to 18
            titleLabel.ForeColor = Color.White;  // Pure white
            titleLabel.Location = new Point(85, 20);
            titleLabel.AutoSize = false;
            titleLabel.Size = new Size(700, 40);
            titleLabel.BackColor = Color.Transparent;
            titleLabel.TextAlign = ContentAlignment.MiddleLeft;
            headerPanel.Controls.Add(titleLabel);
            
            // NO ZOOM LABEL - Removed completely
            
            // Info panel
            Panel infoPanel = new Panel();
            infoPanel.Dock = DockStyle.Top;
            infoPanel.Height = 35;
            infoPanel.BackColor = Color.FromArgb(250, 250, 250);
            
            Label infoLabel = new Label();
            infoLabel.Text = string.Format("Type: {0} | Double-click to fit | Mouse drag to pan | Scroll to zoom", view.ViewType);
            infoLabel.Dock = DockStyle.Fill;
            infoLabel.TextAlign = ContentAlignment.MiddleCenter;
            infoLabel.Font = new Font("Segoe UI", 10);  // Slightly bigger
            infoLabel.ForeColor = Color.FromArgb(80, 80, 80);
            infoPanel.Controls.Add(infoLabel);
            
            // Picture box container
            Panel containerPanel = new Panel();
            containerPanel.Dock = DockStyle.Fill;
            containerPanel.BackColor = Color.White;
            containerPanel.Padding = new Padding(10);
            
            // Zoomable picture box
            pictureBox = new ZoomablePictureBox();
            pictureBox.Dock = DockStyle.Fill;
            pictureBox.BackColor = Color.White;
            
            containerPanel.Controls.Add(pictureBox);
            
            // Add controls
            this.Controls.Add(containerPanel);
            this.Controls.Add(infoPanel);
            this.Controls.Add(headerPanel);
        }
        
        private void HeaderPanel_Paint(object sender, PaintEventArgs e)
        {
            Panel panel = sender as Panel;
            if (panel != null)
            {
                // Nice gradient blue background
                using (var brush = new LinearGradientBrush(
                    panel.ClientRectangle,
                    Color.FromArgb(0, 122, 204),
                    Color.FromArgb(0, 88, 156),
                    LinearGradientMode.Horizontal))
                {
                    e.Graphics.FillRectangle(brush, panel.ClientRectangle);
                }
            }
        }
        
        private void LoadViewImage()
        {
            try
            {
                string tempPath = Path.GetTempPath();
                string fileName = string.Format("ViewPreview_{0}_{1}", view.Id.Value, DateTime.Now.Ticks);
                string filePath = Path.Combine(tempPath, fileName);
                
                // Use DWF export for Revit 2024
                DWFExportOptions options = new DWFExportOptions();
                options.ExportingAreas = false;
                options.MergedViews = false;
                
                System.Collections.Generic.List<ElementId> views = new System.Collections.Generic.List<ElementId>();
                views.Add(view.Id);
                
                Transaction t = new Transaction(doc, "Export View");
                t.Start();
                
                bool exported = false;
                try
                {
                    // Try to export as image using print
                    PrintManager pm = doc.PrintManager;
                    pm.PrintRange = PrintRange.Select;
                    pm.ViewSheetSetting.CurrentViewSheetSet.Views = views;
                    pm.PrintToFile = true;
                    pm.PrintToFileName = filePath + ".png";
                    pm.Apply();
                    exported = true;
                }
                catch
                {
                    // Fall back to simple screen capture
                }
                
                t.RollBack();
                
                if (false) // disabled for now
                {
                    System.Threading.Thread.Sleep(100);
                    
                    // Look for exported file
                    string expectedFile = filePath + ".png";
                    if (!File.Exists(expectedFile))
                    {
                        // Sometimes Revit adds the view name
                        string[] files = System.IO.Directory.GetFiles(tempPath, fileName + "*.png");
                        if (files.Length > 0)
                            expectedFile = files[0];
                    }
                    
                    if (File.Exists(expectedFile))
                    {
                        using (var stream = new FileStream(expectedFile, FileMode.Open, FileAccess.Read))
                        {
                            Image img = Image.FromStream(stream);
                            pictureBox.SetImageAndFit(img);
                        }
                        
                        // Clean up temp file
                        try { File.Delete(expectedFile); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading preview: " + ex.Message);
            }
        }
        
        private string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;
            return text.Substring(0, maxLength - 3) + "...";
        }
    }
    
    // Simple ZoomablePictureBox without zoom label
    public class ZoomablePictureBox : Control
    {
        public ZoomablePictureBox()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
        }
        
        private Image _image;
        private float _zoomFactor = 1.0f;
        private Point _imageLocation;
        private Point _lastMousePos;
        private bool _isPanning;
        
        public void SetImageAndFit(Image image)
        {
            _image = image;
            if (_image != null)
            {
                FitToWindow();
            }
            Invalidate();
        }
        
        protected override void OnPaint(PaintEventArgs e)
        {
            if (_image == null) return;
            
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
            
            int width = (int)(_image.Width * _zoomFactor);
            int height = (int)(_image.Height * _zoomFactor);
            
            e.Graphics.DrawImage(_image, _imageLocation.X, _imageLocation.Y, width, height);
        }
        
        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if (_image == null) return;
            
            float delta = e.Delta > 0 ? 1.1f : 0.9f;
            _zoomFactor *= delta;
            _zoomFactor = Math.Max(0.1f, Math.Min(10f, _zoomFactor));
            
            // Center zoom on mouse position
            int dx = (int)((e.X - _imageLocation.X) * (delta - 1));
            int dy = (int)((e.Y - _imageLocation.Y) * (delta - 1));
            _imageLocation.X -= dx;
            _imageLocation.Y -= dy;
            
            Invalidate();
        }
        
        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _isPanning = true;
                _lastMousePos = e.Location;
                Cursor = Cursors.Hand;
            }
        }
        
        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_isPanning && _image != null)
            {
                int dx = e.X - _lastMousePos.X;
                int dy = e.Y - _lastMousePos.Y;
                _imageLocation.X += dx;
                _imageLocation.Y += dy;
                _lastMousePos = e.Location;
                Invalidate();
            }
        }
        
        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _isPanning = false;
                Cursor = Cursors.Default;
            }
        }
        
        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            if (_image != null)
            {
                FitToWindow();
            }
        }
        
        private void FitToWindow()
        {
            if (_image == null || Width == 0 || Height == 0) return;
            
            float xRatio = (float)Width / _image.Width;
            float yRatio = (float)Height / _image.Height;
            _zoomFactor = Math.Min(xRatio, yRatio) * 0.95f;
            
            // Center the image
            int scaledWidth = (int)(_image.Width * _zoomFactor);
            int scaledHeight = (int)(_image.Height * _zoomFactor);
            _imageLocation = new Point((Width - scaledWidth) / 2, (Height - scaledHeight) / 2);
            
            Invalidate();
        }
    }
}