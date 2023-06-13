using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace CTAImport
{
    public class CTAImportDock : DockableControl
    {
        private CTAImport.CTAIMainForm myform = null;

        public CTAImportDock() : base("CTAImportInternal")
        {
            this.DisplayName = "Import CartoonAnimator project";
        }

        public override DockWindowStyle DefaultDockWindowStyle
        {
            get { return DockWindowStyle.Docked; }
        }

        public override Size DefaultFloatingSize
        {
            get { return new Size(640, 480); }
        }

        protected override void OnLoad(EventArgs args)
        {
            myform = new CTAImport.CTAIMainForm(myVegas);
            myform.Dock = DockStyle.Fill;
            this.Controls.Add(myform);
        }

        protected override void OnClosed(EventArgs args)
        {
             base.OnClosed(args);
        }
    }
}


public class CTAImportCCM : ICustomCommandModule
{
    public Vegas myVegas = null;
    CustomCommand CCM = new CustomCommand(CommandCategory.View, "CTAImport");

    public void InitializeModule(Vegas vegas)
    {
        myVegas = vegas;

        CCM.MenuItemName = "Import CTA Project";
    }

    public ICollection GetCustomCommands()
    {
        CCM.MenuPopup += this.HandlePICmdMenuPopup;
        CCM.Invoked += this.HandlePICmdInvoked;
        CustomCommand[] cmds = new CustomCommand[] { CCM };
        return cmds;
    }

    void HandlePICmdMenuPopup(Object sender, EventArgs args)
    {
        CCM.Checked = myVegas.FindDockView("CTAImportInternal");
    }

    void HandlePICmdInvoked(Object sender, EventArgs args)
    {
        if (!myVegas.ActivateDockView("CTAImportInternal"))
        {
            CTAImport.CTAImportDock CCMDock = new CTAImport.CTAImportDock();
            CCMDock.AutoLoadCommand = CCM;
            CCMDock.PersistDockWindowState = true;
            myVegas.LoadDockView(CCMDock);
        }
    }

}



