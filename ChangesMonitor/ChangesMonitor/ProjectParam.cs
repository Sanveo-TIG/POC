using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Revit.SDK.Samples.ChangesMonitor.CS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using TIGUtility;

namespace ChangesMonitor
{
    [Transaction(TransactionMode.Manual)]
    public class ProjectParameterHandler : IExternalEventHandler
    {
        UIDocument uiDoc = null;
        Document doc = null;
        public void Execute(UIApplication UiApp)
        {
            uiDoc = UiApp.ActiveUIDocument;
            doc = uiDoc.Document;
            try
            {
                using Transaction tx = new Transaction(doc);
                tx.Start("Project Parameter");
                string tempfilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                DirectoryInfo di = new DirectoryInfo(tempfilePath);
                string tempfileName = System.IO.Path.Combine(di.FullName, "AutoUpdate_Parameter.txt");
                UiApp.Application.SharedParametersFilename = tempfileName;
                DefinitionFile defFile = UiApp.Application.OpenSharedParameterFile();

                CategorySet catSet = UiApp.Application.Create.NewCategorySet();
                CategorySet catSet_FORCUID = UiApp.Application.Create.NewCategorySet();
                CategorySet catSet_2 = UiApp.Application.Create.NewCategorySet();
                CategorySet catSet_3 = UiApp.Application.Create.NewCategorySet();

                Categories categories = doc.Settings.Categories;
                Categories categories_2 = doc.Settings.Categories;
                Categories categories_3 = doc.Settings.Categories;

                Category conduit
                 = categories.get_Item(
                   BuiltInCategory.OST_Conduit);
                catSet_FORCUID.Insert(conduit);
                if (defFile != null)
                {
                    foreach (DefinitionGroup dG in defFile.Groups)
                    {
                        foreach (Definition def in dG.Definitions)
                        {
                            if (def.Name == "Bend Angle")
                            {
                                Autodesk.Revit.DB.Binding binding = UiApp.Application.Create.NewInstanceBinding(catSet_FORCUID);
                                BindingMap map = (new UIApplication(UiApp.Application)).ActiveUIDocument.Document.ParameterBindings;
                                map.Insert(def, binding, def.Name == "Bend Angle" ? BuiltInParameterGroup.PG_IDENTITY_DATA : def.ParameterGroup);
                            }
                        }
                    }
                }

                tx.Commit();
            }
            catch
            {
            }
        }


        public string GetName()
        {
            return "Revit Addin";
        }
    }
}
