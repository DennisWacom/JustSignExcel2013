using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using FLSIGCTLLib;
using FlSigCaptLib;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new JustSignExcelRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace JustSignExcel2013
{
    [ComVisible(true)]
    public class JustSignExcelRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public JustSignExcelRibbon()
        {
        }

        public Bitmap getSignatureIcon(Office.IRibbonControl control)
        {
            return JustSignExcel2013.Properties.Resources.sign;
        }

        public void CaptureSignature(Office.IRibbonControl control)
        {
            sign();
        }

        public void sign()
        {
            SigCtl sigCtl = new SigCtl();
            DynamicCapture dc = new DynamicCapture();
            DynamicCaptureResult res = dc.Capture(sigCtl, "Name", "Reason", null, null);

            if (res == DynamicCaptureResult.DynCaptOK)
            {
                SigObj sigObj = (SigObj)sigCtl.Signature;
                //sigObj.set_ExtraData("AdditionalData", "C# test: Additional data");

                String filename = System.IO.Path.GetTempFileName();
                try
                {
                    sigObj.RenderBitmap(filename, 400, 200, "image/png", 1.0f, 0x000000, 0xffffff, 5.0f, 5.0f, RBFlags.RenderOutputFilename | RBFlags.RenderColor32BPP | RBFlags.RenderEncodeData | RBFlags.RenderBackgroundTransparent);

                    Range activecell = Globals.ThisAddIn.Application.Selection;

                    Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;

                    Microsoft.Office.Interop.Excel.Shape signature =
                        ws.Shapes.AddPicture(filename, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, activecell.Left, activecell.Top, 400, 200);

                    signature.Placement = XlPlacement.xlMoveAndSize;
                    signature.Height = (float)activecell.Height;
                    signature.Width = (float)activecell.Width;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }

        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("JustSignExcel2013.JustSignExcelRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
