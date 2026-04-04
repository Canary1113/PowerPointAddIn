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
            this.buttonApplyLiquidGlass = this.Factory.CreateRibbonButton();
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
            this.groupLiquidGlass.Items.Add(this.buttonApplyLiquidGlass);
            this.groupLiquidGlass.Items.Add(this.buttonApplyLiquidGlassDebug);
            this.groupLiquidGlass.Label = "PPT Assistant";
            this.groupLiquidGlass.Name = "groupLiquidGlass";
            // 
            // buttonApplyLiquidGlass
            // 
            this.buttonApplyLiquidGlass.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonApplyLiquidGlass.OfficeImageId = "ShapesFillColorPicker";
            this.buttonApplyLiquidGlass.Label = "Apply Liquid Glass";
            this.buttonApplyLiquidGlass.Name = "buttonApplyLiquidGlass";
            this.buttonApplyLiquidGlass.ScreenTip = "Apply a liquid-glass look";
            this.buttonApplyLiquidGlass.ShowImage = true;
            this.buttonApplyLiquidGlass.SuperTip = "Sets selected shapes to slide-background fill, bright outline, glow, and bevel.";
            this.buttonApplyLiquidGlass.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonApplyLiquidGlass_Click);
            // 
            // buttonApplyLiquidGlassDebug
            // 
            this.buttonApplyLiquidGlassDebug.OfficeImageId = "AnimationPreview";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonApplyLiquidGlassDebug;
    }
}
