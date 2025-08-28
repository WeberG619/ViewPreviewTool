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
                if (uidoc == null) return Result.Failed;
                
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count != 1)
                {
                    TaskDialog.Show("View Preview", "Please select a single view to preview.");
                    return Result.Cancelled;
                }
                
                ElementId selectedId = selectedIds.First();
                Element elem = uidoc.Document.GetElement(selectedId);
                Autodesk.Revit.DB.View view = elem as Autodesk.Revit.DB.View;
                
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
                
                // Create new preview
                previewForm = new ViewPreviewForm(uidoc.Document, view);
                previewForm.TopMost = false;
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
        private ZoomablePictureBox pictureBox;
        
        public ViewPreviewForm(Document document, Autodesk.Revit.DB.View v)
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
            this.TopMost = false;
            
            // Header panel - BIGGER and BETTER
            System.Windows.Forms.Panel headerPanel = new System.Windows.Forms.Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 80;  // Bigger header
            headerPanel.Paint += new PaintEventHandler(HeaderPanel_Paint);
            
            // Logo
            try
            {
                string logoPath = @"D:\BIM_Ops_Studio\BIM_Ops_Studio_logo_all_sizes\BIM_Ops_Studio_logo_1500x1500.png";
                if (File.Exists(logoPath))
                {
                    PictureBox logoPic = new PictureBox();
                    using (var originalImage = Image.FromFile(logoPath))
                    {
                        var logo = new Bitmap(60, 60);
                        using (var g = Graphics.FromImage(logo))
                        {
                            g.SmoothingMode = SmoothingMode.AntiAlias;
                            g.Clear(System.Drawing.Color.Transparent);
                            g.DrawImage(originalImage, 0, 0, 60, 60);
                        }
                        logoPic.Image = logo;
                    }
                    logoPic.SizeMode = PictureBoxSizeMode.Zoom;
                    logoPic.Size = new Size(60, 60);
                    logoPic.Location = new System.Drawing.Point(15, 10);
                    logoPic.BackColor = System.Drawing.Color.Transparent;
                    headerPanel.Controls.Add(logoPic);
                }
            }
            catch { }
            
            // Title - BIGGER and WHITER
            Label titleLabel = new Label();
            titleLabel.Text = TruncateText(view.Name, 50);
            titleLabel.Font = new Font("Segoe UI", 18, FontStyle.Bold);
            titleLabel.ForeColor = System.Drawing.Color.White;
            titleLabel.Location = new System.Drawing.Point(85, 20);
            titleLabel.AutoSize = false;
            titleLabel.Size = new Size(700, 40);
            titleLabel.BackColor = System.Drawing.Color.Transparent;
            titleLabel.TextAlign = ContentAlignment.MiddleLeft;
            headerPanel.Controls.Add(titleLabel);
            
            // NO ZOOM LABEL - Completely removed
            
            // Info panel
            System.Windows.Forms.Panel infoPanel = new System.Windows.Forms.Panel();
            infoPanel.Dock = DockStyle.Top;
            infoPanel.Height = 35;
            infoPanel.BackColor = System.Drawing.Color.FromArgb(250, 250, 250);
            
            Label infoLabel = new Label();
            infoLabel.Text = string.Format("Type: {0} | Double-click to fit | Mouse drag to pan | Scroll to zoom", view.ViewType);
            infoLabel.Dock = DockStyle.Fill;
            infoLabel.TextAlign = ContentAlignment.MiddleCenter;
            infoLabel.Font = new Font("Segoe UI", 10);
            infoLabel.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80);
            infoPanel.Controls.Add(infoLabel);
            
            // Picture box container
            System.Windows.Forms.Panel containerPanel = new System.Windows.Forms.Panel();
            containerPanel.Dock = DockStyle.Fill;
            containerPanel.BackColor = System.Drawing.Color.White;
            containerPanel.Padding = new Padding(10);
            
            // Zoomable picture box
            pictureBox = new ZoomablePictureBox();
            pictureBox.Dock = DockStyle.Fill;
            pictureBox.BackColor = System.Drawing.Color.White;
            
            containerPanel.Controls.Add(pictureBox);
            
            // Add controls
            this.Controls.Add(containerPanel);
            this.Controls.Add(infoPanel);
            this.Controls.Add(headerPanel);
        }
        
        private void HeaderPanel_Paint(object sender, PaintEventArgs e)
        {
            System.Windows.Forms.Panel panel = sender as System.Windows.Forms.Panel;
            if (panel != null)
            {
                using (var brush = new LinearGradientBrush(
                    panel.ClientRectangle,
                    System.Drawing.Color.FromArgb(0, 122, 204),
                    System.Drawing.Color.FromArgb(0, 88, 156),
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
                // Create a simple image for now
                Bitmap bmp = new Bitmap(800, 600);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(System.Drawing.Color.White);
                    g.DrawString(string.Format("View: {0}", view.Name), 
                        new Font("Arial", 24), 
                        Brushes.Black, 
                        new RectangleF(0, 0, 800, 600),
                        new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                }
                
                pictureBox.SetImageAndFit(bmp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        
        private string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;
            return text.Substring(0, maxLength - 3) + "...";
        }
    }
    
    public class ZoomablePictureBox : System.Windows.Forms.Control
    {
        private Image _image;
        private float _zoomFactor = 1.0f;
        private System.Drawing.Point _imageLocation;
        private System.Drawing.Point _lastMousePos;
        private bool _isPanning;
        
        public ZoomablePictureBox()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
        }
        
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
            
            int scaledWidth = (int)(_image.Width * _zoomFactor);
            int scaledHeight = (int)(_image.Height * _zoomFactor);
            _imageLocation = new System.Drawing.Point((Width - scaledWidth) / 2, (Height - scaledHeight) / 2);
            
            Invalidate();
        }
    }
}