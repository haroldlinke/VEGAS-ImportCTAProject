using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace CTAImportTest
{
    public class CCTestDock : DockableControl
    {
        private CTAImportTest.CCMainForm myform = null;

        public CCTestDock() : base("CCTestInternal")
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
            myform = new CTAImportTest.CCMainForm(myVegas);
            myform.Dock = DockStyle.Fill;
            this.Controls.Add(myform);
        }

        protected override void OnClosed(EventArgs args)
        {
             base.OnClosed(args);
        }
    }
}


public class CCTestCCM : ICustomCommandModule
{
    public Vegas myVegas = null;
    CustomCommand CCM = new CustomCommand(CommandCategory.View, "CTAImportTest");

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
        CCM.Checked = myVegas.FindDockView("CTATestInternal");
    }

    void HandlePICmdInvoked(Object sender, EventArgs args)
    {
        if (!myVegas.ActivateDockView("CTATestInternal"))
        {
            CTAImportTest.CCTestDock CCMDock = new CTAImportTest.CCTestDock();
            CCMDock.AutoLoadCommand = CCM;
            CCMDock.PersistDockWindowState = true;
            myVegas.LoadDockView(CCMDock);
        }
    }

}



