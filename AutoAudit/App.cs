using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace AutoAudit
{
    /// <summary>
    /// Simple interface to ensure this app is available 
    /// when Revit is in a zero-document state.
    /// </summary>
    public class Availability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication a, CategorySet b)
        {
            return true;
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class App : IExternalApplication
    {
        /// <summary>
        /// Executes when Revit has started, and is
        /// in a zero-document state.
        /// </summary>
        /// <param name="a"></param>
        /// <returns></returns>
        public Result OnStartup(UIControlledApplication a)
        {
            string tabName = "AMA Tools";
            try
            {
                a.CreateRibbonTab(tabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // Do nothing.
            }
            // Add a new ribbon panel
            RibbonPanel newPanel = a.CreateRibbonPanel(tabName, "AutoAudit");
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData button1Data = new PushButtonData("command",
                "Audit", thisAssemblyPath, "AutoAudit.Command");
            button1Data.AvailabilityClassName = "AutoAudit.Availability";
            PushButton pushButton1 = newPanel.AddItem(button1Data) as PushButton;
            pushButton1.LargeImage = BmpImageSource(@"AutoAudit.Embedded_Media.large.png");


            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        private ImageSource BmpImageSource(string embeddedPath)
        {
            System.IO.Stream manifestResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(embeddedPath);
            PngBitmapDecoder pngBitmapDecoder = new PngBitmapDecoder(manifestResourceStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            return pngBitmapDecoder.Frames[0];
        }
    }
}
