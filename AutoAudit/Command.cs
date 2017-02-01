using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace AutoAudit
{
    [Transaction(TransactionMode.Manual)]
    class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            MainWindow mWindow = new MainWindow(commandData);
            try
            {
                mWindow.ShowDialog();
            }
            catch (Exception)
            {
                TaskDialog.Show("AutoAudit", "File could not be read!");
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}
