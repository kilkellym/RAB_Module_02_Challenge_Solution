#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace RAB_Module_02_Challenge_Solution
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. prompt user to select elements
            TaskDialog.Show("Select Lines", "Select some line to convert to Revit elements.");
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some elements");

            // 2. filter selected elements
            List<CurveElement> filteredList = new List<CurveElement>();
            foreach (Element element in pickList)
            {
                if (element is CurveElement)
                {
                    CurveElement curve = element as CurveElement;
                    //CurveElement curve = (CurveElement) element;

                    filteredList.Add(curve);
                }
            }

            TaskDialog.Show("Curves", $"You selected {filteredList.Count} lines.");
            //TaskDialog.Show("Curves", "You selected" + filteredList.Count.ToString() + " lines.");

            // 3. Get level 
            Level myLevel = GetLevelByName(doc, "Level 1");

            // 4. Get types
            Element wallType1 = GetWallTypeByName(doc, "Storefront");
            Element wallType2 = GetWallTypeByName(doc, "Generic - 8\"");

            Element ductSystemType = GetMEPSystemTypeByName(doc, "Supply Air");
            Element ductType = GetDuctTypeByName(doc, "Default");

            Element pipeSystemType = GetMEPSystemTypeByName(doc, "Domestic Hot Water");
            Element pipeType = GetPipeTypeByName(doc, "Default");

            // 10. create list of elements to delete
            List<ElementId> linesToHide = new List<ElementId>();

            // 5. Loop through selected CurveElements
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Elements");
                foreach (CurveElement currentCurve in filteredList)
                {
                    // 6. Get GraphicStyle and Curve for each CurveElement
                    Curve elementCurve = currentCurve.GeometryCurve;

                    //Curve elementCurve = currentCurve.GeometryCurve;
                    GraphicsStyle currentStyle = currentCurve.LineStyle as GraphicsStyle;

                    // 7. skip arcs and circle
                    if (elementCurve.IsBound == false)
                    {
                        linesToHide.Add(currentCurve.Id);
                        continue;
                    }

                    // 8. get start and points
                    XYZ startPoint = elementCurve.GetEndPoint(0);
                    XYZ endPoint = elementCurve.GetEndPoint(1);

                    // 9. Use Switch statement to create walls, ducts, and pipes
                    switch (currentStyle.Name)
                    {
                        case "A-GLAZ":
                            Wall currentWall = Wall.Create(doc, elementCurve, wallType1.Id,
                                myLevel.Id, 20, 0, false, false);
                            break;

                        case "A-WALL":
                            Wall currentWall2 = Wall.Create(doc, elementCurve, wallType2.Id,
                                myLevel.Id, 20, 0, false, false);
                            break;

                        case "M-DUCT":
                            Duct currentDuct = Duct.Create(doc, ductSystemType.Id,
                                ductType.Id, myLevel.Id, startPoint, endPoint);
                            break;

                        case "P-PIPE":
                            Pipe currentPipe = Pipe.Create(doc, pipeSystemType.Id,
                                pipeType.Id, myLevel.Id, startPoint, endPoint);
                            break;

                        default:
                            linesToHide.Add(currentCurve.Id);
                            break;
                    }
                }

                // 11. hide elements
                doc.ActiveView.HideElements(linesToHide);

                t.Commit();

                return Result.Succeeded;
            }
        }

        private static Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector levelCollector = new FilteredElementCollector(doc);
            levelCollector.OfClass(typeof(Level));
            levelCollector.WhereElementIsNotElementType();

            foreach (Level curLevel in levelCollector)
            {
                if (curLevel.Name == levelName)
                {
                    return curLevel;
                }
            }

            return null;
        }
        private static Element GetWallTypeByName(Document doc, string typeName)
        { 
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (Element curType in collector)
            {
                if (curType.Name == typeName)
                {
                    return curType;
                }
            }

            return null;
        }
        private static Element GetMEPSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element curType in collector)
            {
                if (curType.Name == typeName)
                {
                    return curType;
                }
            }

            return null;
        }
        private static Element GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element curType in collector)
            {
                if (curType.Name == typeName)
                {
                    return curType;
                }
            }

            return null;
        }
        private static Element GetDuctTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));

            foreach (Element curType in collector)
            {
                if (curType.Name == typeName)
                {
                    return curType;
                }
            }

            return null;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
