namespace PowerPointAddIn.UI
{
    partial class PPTAssistantRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public PPTAssistantRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tabAddIns = this.Factory.CreateRibbonTab();
            this.groupBackground = this.Factory.CreateRibbonGroup();
            this.buttonBackground = this.Factory.CreateRibbonButton();
            this.groupGlass = this.Factory.CreateRibbonGroup();
            this.buttonApplyLiquidGlass = this.Factory.CreateRibbonButton();
            this.buttonGlassWhite = this.Factory.CreateRibbonButton();
            this.buttonGlassBlack = this.Factory.CreateRibbonButton();
            this.buttonRecover = this.Factory.CreateRibbonButton();
            this.groupAnimation = this.Factory.CreateRibbonGroup();
            this.buttonNavbarAnimation = this.Factory.CreateRibbonButton();
            this.groupTheme = this.Factory.CreateRibbonGroup();
            this.galleryTheme = this.Factory.CreateRibbonGallery();
            this.buttonThemeColor = this.Factory.CreateRibbonButton();
            this.buttonOpenThemeTemplate = this.Factory.CreateRibbonButton();
            this.tabAddIns.SuspendLayout();
            this.groupBackground.SuspendLayout();
            this.groupGlass.SuspendLayout();
            this.groupAnimation.SuspendLayout();
            this.groupTheme.SuspendLayout();
            this.SuspendLayout();
            //
            // tabAddIns
            //
            this.tabAddIns.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabAddIns.ControlId.OfficeId = "TabAddIns";
            this.tabAddIns.Groups.Add(this.groupBackground);
            this.tabAddIns.Groups.Add(this.groupGlass);
            this.tabAddIns.Groups.Add(this.groupAnimation);
            this.tabAddIns.Groups.Add(this.groupTheme);
            this.tabAddIns.Name = "tabAddIns";
            //
            // groupBackground
            //
            this.groupBackground.Items.Add(this.buttonBackground);
            this.groupBackground.Label = "Background";
            this.groupBackground.Name = "groupBackground";
            //
            // buttonBackground
            //
            this.buttonBackground.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonBackground.OfficeImageId = "PictureInsertFromFile";
            this.buttonBackground.Label = "Background Blur";
            this.buttonBackground.Name = "buttonBackground";
            this.buttonBackground.ScreenTip = "Fill the slide with the selected picture";
            this.buttonBackground.ShowImage = true;
            this.buttonBackground.SuperTip = "Adds a sharp full-slide picture and sets a blurred copy as the slide background.";
            this.buttonBackground.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonBackground_Click);
            //
            // groupGlass
            //
            this.groupGlass.Items.Add(this.buttonRecover);
            this.groupGlass.Items.Add(this.buttonApplyLiquidGlass);
            this.groupGlass.Items.Add(this.buttonGlassWhite);
            this.groupGlass.Items.Add(this.buttonGlassBlack);
            this.groupGlass.Label = "Glass";
            this.groupGlass.Name = "groupGlass";
            //
            // buttonApplyLiquidGlass
            //
            this.buttonApplyLiquidGlass.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonApplyLiquidGlass.OfficeImageId = "ShapeEffectsMenu";
            this.buttonApplyLiquidGlass.Label = "Transparent Glass";
            this.buttonApplyLiquidGlass.Name = "buttonApplyLiquidGlass";
            this.buttonApplyLiquidGlass.ScreenTip = "Apply a transparent glass look";
            this.buttonApplyLiquidGlass.ShowImage = true;
            this.buttonApplyLiquidGlass.SuperTip = "Applies a transparent glass look with background fill, bevel, and shadow to the selected shape.";
            this.buttonApplyLiquidGlass.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonApplyLiquidGlass_Click);
            //
            // buttonGlassWhite
            //
            this.buttonGlassWhite.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonGlassWhite.OfficeImageId = "ColorWhite";
            this.buttonGlassWhite.Label = "White Glass";
            this.buttonGlassWhite.Name = "buttonGlassWhite";
            this.buttonGlassWhite.ScreenTip = "Create a white glass overlay";
            this.buttonGlassWhite.ShowImage = true;
            this.buttonGlassWhite.SuperTip = "Duplicates the selected shape, overlays semi-transparent white, and groups the two shapes.";
            this.buttonGlassWhite.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGlassWhite_Click);
            //
            // buttonGlassBlack
            //
            this.buttonGlassBlack.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonGlassBlack.OfficeImageId = "ColorBlack";
            this.buttonGlassBlack.Label = "Black Glass";
            this.buttonGlassBlack.Name = "buttonGlassBlack";
            this.buttonGlassBlack.ScreenTip = "Create a black glass overlay";
            this.buttonGlassBlack.ShowImage = true;
            this.buttonGlassBlack.SuperTip = "Duplicates the selected shape, overlays semi-transparent black, and groups the two shapes.";
            this.buttonGlassBlack.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGlassBlack_Click);
            //
            // buttonRecover
            //
            this.buttonRecover.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRecover.Label = "Recover";
            this.buttonRecover.Name = "buttonRecover";
            this.buttonRecover.OfficeImageId = "Undo";
            this.buttonRecover.ScreenTip = "Remove Glass effects";
            this.buttonRecover.ShowImage = true;
            this.buttonRecover.SuperTip = "Removes Glass effects generated by this add-in and restores the base shape as much as possible.";
            this.buttonRecover.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRecover_Click);
            //
            // groupAnimation
            //
            this.groupAnimation.Items.Add(this.buttonNavbarAnimation);
            this.groupAnimation.Label = "Animation";
            this.groupAnimation.Name = "groupAnimation";
            //
            // buttonNavbarAnimation
            //
            this.buttonNavbarAnimation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonNavbarAnimation.Label = "Nav Bar Animation";
            this.buttonNavbarAnimation.Name = "buttonNavbarAnimation";
            this.buttonNavbarAnimation.OfficeImageId = "AnimationPreview";
            this.buttonNavbarAnimation.ScreenTip = "Apply the navbar move-and-scale animation";
            this.buttonNavbarAnimation.ShowImage = true;
            this.buttonNavbarAnimation.SuperTip = "Adds a 1-second straight move, a simultaneous 0.5-second grow to 125%, and a delayed 0.5-second shrink to 80%.";
            this.buttonNavbarAnimation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNavbarAnimation_Click);
            //
            // groupTheme
            //
            this.groupTheme.Items.Add(this.galleryTheme);
            this.groupTheme.Items.Add(this.buttonThemeColor);
            this.groupTheme.Items.Add(this.buttonOpenThemeTemplate);
            this.groupTheme.Label = "Theme";
            this.groupTheme.Name = "groupTheme";
            //
            // galleryTheme
            //
            this.galleryTheme.ColumnCount = 1;
            this.galleryTheme.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.galleryTheme.Label = "Theme";
            this.galleryTheme.Name = "galleryTheme";
            this.galleryTheme.OfficeImageId = "SlideDesign";
            this.galleryTheme.RowCount = 8;
            this.galleryTheme.ScreenTip = "Apply a slide theme";
            this.galleryTheme.ShowImage = true;
            this.galleryTheme.ShowItemImage = false;
            this.galleryTheme.ShowItemLabel = true;
            this.galleryTheme.SuperTip = "Applies a template-based light or dark slide theme.";
            this.galleryTheme.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryTheme_Click);
            this.galleryTheme.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.galleryTheme_ItemsLoading);
            //
            // buttonThemeColor
            //
            this.buttonThemeColor.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonThemeColor.Label = "Color";
            this.buttonThemeColor.Name = "buttonThemeColor";
            this.buttonThemeColor.OfficeImageId = "FontColorPicker";
            this.buttonThemeColor.ScreenTip = "Customize theme color";
            this.buttonThemeColor.ShowImage = true;
            this.buttonThemeColor.SuperTip = "Choose a color to replace the theme accent text and line color on the current page.";
            this.buttonThemeColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonThemeColor_Click);
            //
            // buttonOpenThemeTemplate
            //
            this.buttonOpenThemeTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonOpenThemeTemplate.Label = "Open Template";
            this.buttonOpenThemeTemplate.Name = "buttonOpenThemeTemplate";
            this.buttonOpenThemeTemplate.OfficeImageId = "FileOpen";
            this.buttonOpenThemeTemplate.ScreenTip = "Open the theme template";
            this.buttonOpenThemeTemplate.ShowImage = true;
            this.buttonOpenThemeTemplate.SuperTip = "Opens the PowerPoint template used to populate the Theme menu.";
            this.buttonOpenThemeTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonOpenThemeTemplate_Click);
            //
            // PPTAssistantRibbon
            //
            this.Name = "PPTAssistantRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabAddIns);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PPTAssistantRibbon_Load);
            this.tabAddIns.ResumeLayout(false);
            this.tabAddIns.PerformLayout();
            this.groupBackground.ResumeLayout(false);
            this.groupBackground.PerformLayout();
            this.groupGlass.ResumeLayout(false);
            this.groupGlass.PerformLayout();
            this.groupAnimation.ResumeLayout(false);
            this.groupAnimation.PerformLayout();
            this.groupTheme.ResumeLayout(false);
            this.groupTheme.PerformLayout();
            this.ResumeLayout(false);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAddIns;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupBackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupGlass;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnimation;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTheme;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonApplyLiquidGlass;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGlassWhite;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGlassBlack;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonBackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRecover;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNavbarAnimation;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryTheme;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonThemeColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonOpenThemeTemplate;
    }
}
