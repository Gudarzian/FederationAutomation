using System;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace FederationAutomation.NavisworksVisualStyle
{
    [Plugin("FederationAutomationVisualStyle", "FEDAUTO", DisplayName = "Federation Automation Visual Style")]
    [AddInPlugin(AddInLocation.None)]
    public sealed class NavisworksVisualStylePlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (document == null || document.IsClear)
                {
                    return 1;
                }

                Viewpoint viewpoint = document.CurrentViewpoint.CreateCopy();
                viewpoint.RenderStyle = ViewpointRenderStyle.FullRender;
                document.CurrentViewpoint.CopyFrom(viewpoint);

                // Keep the standard Navisworks graduated sky background in the saved document.
                document.SetGraduatedBackground(
                    Color.FromByteRGB(128, 179, 232),
                    Color.FromByteRGB(235, 245, 255));
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("Federation Automation could not apply the Navisworks visual style: " + exception.Message);
                return 1;
            }
        }
    }
}
