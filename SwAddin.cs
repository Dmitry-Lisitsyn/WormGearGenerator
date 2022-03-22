using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using WormGearGenerator.Properties;
using Xarial.XCad.Base.Attributes;
using Xarial.XCad.SolidWorks;
using Xarial.XCad.SolidWorks.UI.PropertyPage;
using Xarial.XCad.UI.Commands;
using Xarial.XCad.UI.Commands.Attributes;
using Xarial.XCad.UI.Commands.Enums;
using Xarial.XCad.UI.PropertyPage.Attributes;
using WormGearGenerator;

namespace WormGearGenerator
{
    [ComVisible(true)]
    [Guid("B5AB2AC9-B4B7-4826-A58C-CC774FADE641")]
    [DisplayName("Worm Gear Generator")]
    [Description("Description of addin")]
    public class SwAddin : SwAddInEx
    {
        public MainWindow window;
        [Title("Worm Gear Generator")]
        [Description("Description of addin")]
        [Icon(typeof(Resources), nameof(Resources.Icon))]

        [ComRegisterFunction]
        public new static void RegisterFunction(Type t)
        {
            SwAddInEx.RegisterFunction(t);
        }

        [ComUnregisterFunction]
        public new static void UnregisterFunction(Type t)
        {
            SwAddInEx.UnregisterFunction(t);
        }

        private enum WormGear
        {
           [Title("Worm Gear Generator")]
           [Description("Description of addin")]
           [Icon(typeof(Resources), nameof(Resources.Icon))]
           [CommandItemInfo(true, true, WorkspaceTypes_e.Part, true)]
            CreateWpfForm
        }

        public override void OnConnect()
        {
            CommandManager.AddCommandGroup<WormGear>().CommandClick += OnCommandClick;
        }

        private void OnCommandClick(WormGear cmd)
        {
            var activeDoc = Application.Documents.Active;

            switch (cmd)
            {
                case WormGear.CreateWpfForm:
                    var wpfPopupWnd = CreatePopupWindow<MainWindow>();
                    wpfPopupWnd.ShowDialog();
                    break;
            }
        }

        public override void OnDisconnect()
        {
            this.Dispose();
        }

    }
}