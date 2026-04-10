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
            this.groupLiquidGlass = this.Factory.CreateRibbonGroup();
            this.buttonBackground = this.Factory.CreateRibbonButton();
            this.buttonApplyLiquidGlass = this.Factory.CreateRibbonButton();
            this.buttonGlassWhite = this.Factory.CreateRibbonButton();
            this.buttonGlassBlack = this.Factory.CreateRibbonButton();
            this.buttonRecover = this.Factory.CreateRibbonButton();
            this.buttonApplyLiquidGlassDebug = this.Factory.CreateRibbonButton();
            this.tabAddIns.SuspendLayout();
            this.groupLiquidGlass.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabAddIns
            // 
            this.tabAddIns.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabAddIns.ControlId.OfficeId = "TabAddIns";
            this.tabAddIns.Groups.Add(this.groupLiquidGlass);
            this.tabAddIns.Name = "tabAddIns";
            // 
            // groupLiquidGlass
            // 
            this.groupLiquidGlass.Items.Add(this.buttonBackground);
            this.groupLiquidGlass.Items.Add(this.buttonApplyLiquidGlass);
            this.groupLiquidGlass.Items.Add(this.buttonGlassWhite);
            this.groupLiquidGlass.Items.Add(this.buttonGlassBlack);
            this.groupLiquidGlass.Items.Add(this.buttonRecover);
            this.groupLiquidGlass.Items.Add(this.buttonApplyLiquidGlassDebug);
            this.groupLiquidGlass.Label = "Add-In";
            this.groupLiquidGlass.Name = "groupLiquidGlass";
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
            this.buttonRecover.Label = "Recover";
            this.buttonRecover.Name = "buttonRecover";
            this.buttonRecover.OfficeImageId = "Undo";
            this.buttonRecover.ScreenTip = "Remove Glass effects";
            this.buttonRecover.ShowImage = true;
            this.buttonRecover.SuperTip = "Removes Glass effects generated by this add-in and restores the base shape as much as possible.";
            this.buttonRecover.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRecover_Click);
            // 
            // buttonApplyLiquidGlassDebug
            // 
            this.buttonApplyLiquidGlassDebug.OfficeImageId = "ReviewNewComment";
            this.buttonApplyLiquidGlassDebug.Label = "Debug Apply";
            this.buttonApplyLiquidGlassDebug.Name = "buttonApplyLiquidGlassDebug";
            this.buttonApplyLiquidGlassDebug.ScreenTip = "Apply with debug prompts";
            this.buttonApplyLiquidGlassDebug.ShowImage = true;
            this.buttonApplyLiquidGlassDebug.SuperTip = "Shows message boxes between key steps to help locate freezes or crashes.";
            this.buttonApplyLiquidGlassDebug.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonApplyLiquidGlassDebug_Click);
            // 
            // PPTAssistantRibbon
            // 
            this.Name = "PPTAssistantRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabAddIns);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PPTAssistantRibbon_Load);
            this.tabAddIns.ResumeLayout(false);
            this.tabAddIns.PerformLayout();
            this.groupLiquidGlass.ResumeLayout(false);
            this.groupLiquidGlass.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAddIns;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupLiquidGlass;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonApplyLiquidGlass;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGlassWhite;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGlassBlack;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonBackground;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRecover;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonApplyLiquidGlassDebug;
    }
}
