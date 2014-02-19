using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using UpdateChecker;

namespace swsp
{
    public class SWIntegration : ISwAddin
    {
        public SldWorks swApp;
        private int swCookie;

        public SWIntegration()
        {
        }

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = (SldWorks)ThisSW;
            swCookie = Cookie;
            // Set-up add-in call back info
            bool result = swApp.SetAddinCallbackInfo(0, this, Cookie);
            this.UISetup();
            return true;
        }
        public bool DisconnectFromSW()
        {
            this.UITeardown();
            return true;
        }
        private void UISetup()
        {
            try
            {
                // Set current UI locale to user settings, not windows language version
                locale.Culture = Thread.CurrentThread.CurrentCulture;
                
                // For debug
                //locale.Culture = new System.Globalization.CultureInfo("en");

                // Clean up first
                this.UITeardown();

                // Build the UI
                Assembly thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());
                CommandManager iCmdMgr = swApp.GetCommandManager(swCookie);
                ICommandGroup cmdGroup = default(ICommandGroup);
                string Title = "Standard Primitives";
                string ToolTip = "Standard primitives";
                int[] docTypes = { (int)swDocumentTypes_e.swDocPART };
                cmdGroup = iCmdMgr.CreateCommandGroup(9934, Title, ToolTip, "", -1);
                // Set up icon list files
                cmdGroup.LargeIconList = Path.Combine(GetAssemblyLocation(), @"icons\icons_16.png");
                cmdGroup.SmallIconList = Path.Combine(GetAssemblyLocation(), @"icons\icons_16.png");
                // we store index of every button to use later
                int[] cmdIdx = new int[10];
                string t;
                t = locale.LProfileSketch;
                cmdIdx[0] = cmdGroup.AddCommandItem2(t, -1, "", t, 0, "L_ProfileSketch", "", 0, (int)swCommandItemType_e.swToolbarItem);
                t = locale.UProfileSketch;
                cmdIdx[1] = cmdGroup.AddCommandItem2(t, -1, "", t, 1, "U_ProfileSketch", "", 1, (int)swCommandItemType_e.swToolbarItem);
                t = locale.TProfileSketchClosed;
                cmdIdx[2] = cmdGroup.AddCommandItem2(t, -1, "", t, 3, "T_ProfileClosedSketch", "", 2, (int)swCommandItemType_e.swToolbarItem);
                t = locale.HexagonSketch;
                cmdIdx[3] = cmdGroup.AddCommandItem2(t, -1, "", t, 4, "HexagonSketch", "", 3, (int)swCommandItemType_e.swToolbarItem);
                t = locale.CircleSketch;
                cmdIdx[4] = cmdGroup.AddCommandItem2(t, -1, "", t, 5, "CircleSketch", "", 4, (int)swCommandItemType_e.swToolbarItem);
                t = locale.UpdateCheck;
                cmdIdx[5] = cmdGroup.AddCommandItem2(t, -1, "", t, 6, "UpdateCheck", "", 5, (int)swCommandItemType_e.swToolbarItem);
                t = locale.About;
                cmdIdx[6] = cmdGroup.AddCommandItem2(t, -1, "", t, 7, "AboutForm", "", 6, (int)swCommandItemType_e.swToolbarItem);
                t = locale.RectangleSketch;
                cmdIdx[7] = cmdGroup.AddCommandItem2(t, -1, "", t, 8, "RectangleSketch", "", 7, (int)swCommandItemType_e.swToolbarItem);
                t = locale.RectangleWithRevolveAxisSketch;
                cmdIdx[8] = cmdGroup.AddCommandItem2(t, -1, "", t, 9, "RectangleWithRevolveAxisSketch", "", 8, (int)swCommandItemType_e.swToolbarItem);
                t = locale.CircleWithRevolveAxisSketch;
                cmdIdx[9] = cmdGroup.AddCommandItem2(t, -1, "", t, 10, "CircleWithRevolveAxisSketch", "", 9, (int)swCommandItemType_e.swToolbarItem);
                cmdGroup.HasToolbar = true;
                cmdGroup.HasMenu = false;
                cmdGroup.Activate();
                foreach (int docType in docTypes)
                {
                    CommandTab cmdTab = iCmdMgr.AddCommandTab(docType, Title);
                    // Group 1
                    CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();
                    int[] cmdIDs1 = new int[2];
                    int[] TextType1 = new int[2];
                    cmdIDs1[0] = cmdGroup.get_CommandID(cmdIdx[0]);
                    TextType1[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs1[1] = cmdGroup.get_CommandID(cmdIdx[1]);
                    TextType1[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    // Group 2
                    CommandTabBox cmdBox2 = cmdTab.AddCommandTabBox();
                    int[] cmdIDs2 = new int[6];
                    int[] TextType2 = new int[6];
                    cmdIDs2[0] = cmdGroup.get_CommandID(cmdIdx[2]);
                    TextType2[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs2[1] = cmdGroup.get_CommandID(cmdIdx[3]);
                    TextType2[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs2[2] = cmdGroup.get_CommandID(cmdIdx[4]);
                    TextType2[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs2[3] = cmdGroup.get_CommandID(cmdIdx[7]);
                    TextType2[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs2[4] = cmdGroup.get_CommandID(cmdIdx[8]);
                    TextType2[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs2[5] = cmdGroup.get_CommandID(cmdIdx[9]);
                    TextType2[5] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    // Group 3
                    CommandTabBox cmdBox3 = cmdTab.AddCommandTabBox();
                    int[] cmdIDs3 = new int[2];
                    int[] TextType3 = new int[2];
                    cmdIDs3[0] = cmdGroup.get_CommandID(cmdIdx[5]);
                    TextType3[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    cmdIDs3[1] = cmdGroup.get_CommandID(cmdIdx[6]);
                    TextType3[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    // Add the commands
                    cmdBox1.AddCommands(cmdIDs1, TextType1);
                    cmdBox2.AddCommands(cmdIDs2, TextType2);
                    cmdBox3.AddCommands(cmdIDs3, TextType3);
                    // Add separators
                    cmdTab.AddSeparator(cmdBox2, cmdIDs2[0]);
                    cmdTab.AddSeparator(cmdBox3, cmdIDs3[0]);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                throw e;
            }
        }

        private void UITeardown()
        {
            CommandManager iCmdMgr = swApp.GetCommandManager(swCookie);
            ICommandTab cmdTab = iCmdMgr.GetCommandTab((int)swDocumentTypes_e.swDocPART, "Standard Primitives");
            if (cmdTab != null)
            {
                iCmdMgr.RemoveCommandTab((CommandTab)cmdTab);
            }
            iCmdMgr.RemoveCommandGroup2(9934, false);
        }

        private string GetAssemblyLocation()
        {
            return System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.LastIndexOf(@"\"));
        }

        public void UpdateCheck()
        {
            UpdateChecker.UpdateChecker.start("http://traal.eu/swsp/update.xml", "SolidWorks Standard Primitives", "swsp", false);
        }

        public void AboutForm()
        {
            Form af = new About();
            af.ShowDialog();
        }

        public void L_ProfileSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2;
            line1 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.05, 0.0, 0.0)); // horizontal
            line2 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.0, 0.05, 0.0)); // vertical
            // Add dimensions
            line1.Select4(false, null);
            swDoc.AddDimension2(0.01, -0.01, 0.0);
            line2.Select4(false, null);
            swDoc.AddDimension2(-0.01, 0.01, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void U_ProfileSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2, line3;
            SketchPoint origin;
            line1 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, 0.05, 0.0, 0.0)); // horizontal
            line2 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, -0.05, 0.05, 0.0)); // left
            line3 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.05, 0.0, 0.0, 0.05, 0.05, 0.0)); // right
            // left and right line same length
            line2.Select4(false, null);
            line3.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            // line 1 center on datum origin
            //Feature f = (Feature)swDoc.FirstFeature();
            //while (f != null)
            //{
            //    if (f.GetTypeName2() == "OriginProfileFeature")
            //    {
            //        f.Select2(false, 0);
            //        break;
            //    }
            //    f = (Feature)f.GetNextFeature();
            //}
            //swDoc.Extension.SelectByID2("", swSelectType_e.swSelDATUMPOINTS.ToString(), 0.0, 0.0, 0.0, false, 0, null, 0);
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            line1.Select4(true, null);
            swDoc.SketchAddConstraints("sgATMIDDLE");
            // Add dimensions
            line1.Select4(false, null);
            swDoc.AddDimension2(0.0, -0.05, 0.0);
            line2.Select4(false, null);
            swDoc.AddDimension2(-0.06, 0.025, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void U_ProfileFlangeSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2, line3, line4, line5;
            SketchPoint origin;
            line1 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, 0.05, 0.0, 0.0)); // horizontal
            line2 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, -0.05, 0.05, 0.0)); // left vertical
            line3 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.05, 0.0, 0.0, 0.05, 0.05, 0.0)); // right vertical
            line4 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.1, 0.05, 0.0, -0.05, 0.05, 0.0)); // left horizontal
            line5 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.05, 0.05, 0.0, 0.1, 0.05, 0.0)); // right horizontal
            // left and right lines same length
            line2.Select4(false, null);
            line3.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            line4.Select4(false, null);
            line5.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            // line 1 center on datum origin
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            line1.Select4(true, null);
            swDoc.SketchAddConstraints("sgATMIDDLE");
            // Add dimensions
            line1.Select4(false, null);
            swDoc.AddDimension2(0.0, -0.05, 0.0);
            line2.Select4(false, null);
            swDoc.AddDimension2(-0.06, 0.025, 0.0);
            line4.Select4(false, null);
            swDoc.AddDimension2(-0.09, 0.025, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void T_ProfileClosedSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2, line3, line4, line5, line6, line7, line8, axis;
            SketchPoint origin;
            // Add the origin point
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            // now the lines
            line1 = swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, 0.05, 0.0, 0.0);
            line2 = swDoc.SketchManager.CreateLine(0.05, 0.0, 0.0, 0.05, 0.01, 0.0);
            line3 = swDoc.SketchManager.CreateLine(0.05, 0.01, 0.0, 0.01, 0.01, 0.0);
            line4 = swDoc.SketchManager.CreateLine(0.01, 0.01, 0.0, 0.01, 0.1, 0.0);
            line5 = swDoc.SketchManager.CreateLine(0.01, 0.1, 0.0, -0.01, 0.1, 0.0);
            line6 = swDoc.SketchManager.CreateLine(-0.01, 0.1, 0.0, -0.01, 0.01, 0.0);
            line7 = swDoc.SketchManager.CreateLine(-0.01, 0.01, 0.0, -0.05, 0.01, 0.0);
            line8 = swDoc.SketchManager.CreateLine(-0.05, 0.01, 0.0, -0.05, 0.0, 0.0);
            axis = swDoc.SketchManager.CreateCenterLine(0.0, 0.0, 0.0, 0.0, 0.1, 0.0);
            // add constraints
            axis.Select4(false, null);
            line4.Select4(true, null);
            line6.Select4(true, null);
            swDoc.SketchAddConstraints("sgSYMMETRIC");
            swDoc.ClearSelection2(true);
            line3.Select4(false, null);
            line7.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            swDoc.SketchAddConstraints("sgCOLINEAR");
            // Add dimensions
            swDoc.ClearSelection2(true);
            line1.Select4(false, null);
            swDoc.AddDimension2(0.0, -0.01, 0.0);
            line8.Select4(false, null);
            swDoc.AddDimension2(-0.051, 0.005, 0.0);
            line1.Select4(false, null);
            line5.Select4(true, null);
            swDoc.AddDimension2(0.055, 0.05, 0.0);
            swDoc.ClearSelection2(true);
            line5.Select4(false, null);
            swDoc.AddDimension2(0.0, 0.11, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void HexagonSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            object[] hexagon;
            hexagon = (object[])swDoc.SketchManager.CreatePolygon(0.0, 0.0, 0.0, 0.0, 0.05, 0.0, 6, true);
            swDoc.ClearSelection2(true);
            foreach (object x in hexagon)
            {
                // select one line and make it horizontal
                SketchSegment y = (SketchSegment)x;
                if (y.GetType() == (int)swSketchSegments_e.swSketchLINE)
                {
                    y.Select4(false, null);
                    swDoc.SketchAddConstraints("sgHORIZONTAL2D");
                    break;
                }
            }
            foreach (object x in hexagon)
            {
                // select the circle and add dimension to it
                SketchSegment y = (SketchSegment)x;
                if (y.GetType() == (int)swSketchSegments_e.swSketchARC)
                {
                    y.Select4(false, null);
                    swDoc.AddDimension2(-0.075, 0.075, 0.0);
                    break;
                }
            }
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void CircleSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            SketchSegment circle;
            circle = (SketchSegment)swDoc.SketchManager.CreateCircle(0.0, 0.0, 0.0, 0.05, 0.0, 0.0);
            circle.Select4(false, null);
            swDoc.AddDimension2(0.075, 0.075, 0.0);
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void RectangleSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            object[] rectangle;
            rectangle = (object[])swDoc.SketchManager.CreateCenterRectangle(0.0, 0.0, 0.0, 0.05, 0.025, 0.0);
            SketchSegment l1, l2;
            l1 = (SketchSegment)rectangle[0];
            l2 = (SketchSegment)rectangle[1];
            l1.Select4(false, null);
            swDoc.AddDimension2(0.0, 0.075, 0.0);
            l2.Select4(false, null);
            swDoc.AddDimension2(0.075, 0.0125, 0.0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void RectangleWithRevolveAxisSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            SketchSegment axis, l0, l1, l2, l3;
            object[] rectangle;
            SketchPoint origin;
            // Add the origin point
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            axis = swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.0, 0.05, 0.0);
            axis.ConstructionGeometry = true;
            rectangle = (object[])swDoc.SketchManager.CreateCornerRectangle(0.05, 0.0, 0.0, 0.075, 0.05, 0.0);
            l0 = (SketchSegment)rectangle[0];
            l1 = (SketchSegment)rectangle[1];
            l2 = (SketchSegment)rectangle[2];
            l3 = (SketchSegment)rectangle[3];
            axis.Select4(false, null);
            l1.Select4(true, null);
            swDoc.AddDimension2(0.025, -0.05, 0.0);
            swDoc.ClearSelection2(true);
            l2.Select4(false, null);
            swDoc.AddDimension2(0.0625, 0.05, 0.0);
            swDoc.ClearSelection2(true);
            l3.Select4(false, null);
            swDoc.AddDimension2(0.1, 0.0, 0.0);
            swDoc.ClearSelection2(true);
            l0.Select4(false, null);
            origin.Select4(true, null);
            swDoc.SketchAddConstraints("sgCOINCIDENT");
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void CircleWithRevolveAxisSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)swApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            SketchSegment axis, circle;
            SketchPoint origin, c_origin;
            // Add the origin point
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            swDoc.ClearSelection2(true);
            axis = swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.0, 0.05, 0.0);
            axis.ConstructionGeometry = true;
            c_origin = swDoc.SketchManager.CreatePoint(0.05, 0.0, 0.0);
            c_origin.Select4(false, null);
            origin.Select4(true, null);
            swDoc.SketchAddConstraints("sgHORIZONTALPOINTS2D");
            circle = swDoc.SketchManager.CreateCircleByRadius(0.05, 0.0, 0.0, 0.025);
            axis.Select4(false, null);
            circle.Select4(true, null);
            swDoc.ClearSelection2(true);
            circle.Select4(false, null);
            origin.Select4(true, null);
            swDoc.AddDimension2(0.025, 0.05, 0.0);
            swDoc.ClearSelection2(true);
            circle.Select4(false, null);
            swDoc.AddDimension2(0.075, 0.05, 0.0);
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        #region COM

        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            string keyPath = String.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            using (Microsoft.Win32.RegistryKey rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                rk.SetValue(null, 1); // Load at startup
                rk.SetValue("Title", "Standard Primitives"); // Title
                rk.SetValue("Description", "Standard primitives addon"); // Description
            }
        }
        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            string keyPath = String.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }

        #endregion
    }
}