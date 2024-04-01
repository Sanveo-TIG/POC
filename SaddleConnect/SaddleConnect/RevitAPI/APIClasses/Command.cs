#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using TIGUtility;
#endregion

namespace SaddleConnect
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        /// <summary>
        /// External command mainline
        /// </summary>
        public Result Execute(
                     ExternalCommandData commandData,
                     ref string message,
                     ElementSet elements)
        {
            try
            {

                if (true)//Utility.HasValidLicense(Util.ProductVersion))
                {
                    CustomUIApplication customUIApplication = new CustomUIApplication
                    {
                        CommandData = commandData
                    };
                    System.Windows.Window window = new MainWindow(customUIApplication);
                    window.Show();
                    window.Closed += OnClosing;
                    if (App.SaddleConnectButton != null)
                        App.SaddleConnectButton.Enabled = false;
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
        public void OnClosing(object senSaddleConnectr, EventArgs e)
        {
            if (App.SaddleConnectButton != null)
                App.SaddleConnectButton.Enabled = true;
        }
    }

}
