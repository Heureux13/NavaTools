using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Diagnostics;

[Transaction(TransactionMode.Manual)]
public class OpenWebsiteCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        Process.Start("https://www.procore.com/");
        return Result.Succeeded;
    }
}