using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Colors;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Specialized;
using Autodesk.AutoCAD.PlottingServices;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using System.Collections;
namespace ACADLIBRARY
{
    public class AcadBase
    {
        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int AcedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);
        #region "Các hàm input"
        public static List<Point2d> GetTwoPoint() //Tra ve 2 point bằng cách chọn Polynine (chon hai dinh cheo nhau cua hinh vuong)
        {
            Point2d p1 = new Point2d(), p2 = new Point2d();
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDB = acDoc.Database;
            Editor acEditor = acDoc.Editor;

            // Create a TypedValue array to define the filter criteria
            TypedValue[] acTypValAr = new TypedValue[1];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);

            PromptSelectionResult prSlRls = acEditor.GetSelection(acSelFtr);
            using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
            {
                if (prSlRls.Status == PromptStatus.OK)
                {
                    SelectionSet acSlSet = prSlRls.Value;
                    foreach (SelectedObject acSSObj in acSlSet)
                    {
                        Polyline acEnt = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;
                        p1 = acEnt.GetPoint2dAt(0);
                        p2 = acEnt.GetPoint2dAt(2);
                    }
                }
            }
            List<Point2d> listPoint = new List<Point2d>{p1,p2};
            return listPoint;
        }
        public static Extents2d GetBoundBox(ObjectId[] entIds, Transaction tran)   // trả về khung Extents2d bao ngoài đối tượng được chọn
        {
            Extents3d ex = new Extents3d();
            foreach (var id in entIds)
            {
                Entity ent = tran.GetObject(id, OpenMode.ForRead) as Entity;
                ex = ent.GeometricExtents;
                ex.AddExtents(ent.GeometricExtents);
            }
            Extents2d ketqua = new Extents2d(ex.MinPoint.X, ex.MinPoint.Y, ex.MaxPoint.X, ex.MaxPoint.Y);
            return ketqua;
        }
        public static Extents2d GetExtents2D() //Tra ve khung Extent3d bang selection Point
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDB = acDoc.Database;
            Editor acEditor = acDoc.Editor;
            PromptPointResult prPoint1;
            PromptPointResult prPoint2;
            Point2d point1 ;
            Point2d point2 ;

            PromptPointOptions prPoint1Opt = new PromptPointOptions("Pick firt Point");
            
            using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
            {
                prPoint1 = acEditor.GetPoint(prPoint1Opt);
                if (prPoint1.Status == PromptStatus.OK)
                {
                    prPoint2 = acEditor.GetCorner("Pick second Point", prPoint1.Value);
                    if ((prPoint1.Value.X == prPoint2.Value.X) || (prPoint1.Value.Y == prPoint2.Value.Y)) //Kiểm tra xem hai điểm có tạo thành đường chéo không
                    {
                        acEditor.WriteMessage("Error: Invalid input");
                        point1 = new Point2d(0, 0);
                        point2 = new Point2d(0, 0);
                    }
                    else
                    {
                        point1 = new Point2d(prPoint1.Value.X, prPoint1.Value.Y);
                        point2 = new Point2d(prPoint2.Value.X, prPoint2.Value.Y);
                    }
                }
                else
                {
                    point1 = new Point2d(0, 0);
                    point2 = new Point2d(0, 0);
                }
                
            }
            return new Extents2d(point1, point2);
        }
        public static List<Extents2d> GetWindow1()   // tra ve mot list window tu block hoặc rectange
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDB = acDoc.Database;
            Editor acEditor = acDoc.Editor;
            List<Extents2d> listWindows = new List<Extents2d>();
            Extents3d ex = new Extents3d();
            Point3d p1 = new Point3d(), p2 = new Point3d();

            // Create a TypedValue array to define the filter criteria
            TypedValue[] acTypValAr = new TypedValue[4];
            //TypedValue[] tvs = new TypedValue[] { new TypedValue((int)DxfCode.Operator,"<or"),new TypedValue((int)DxfCode.Operator,"<and"),new TypedValue((int)DxfCode.LayerName,"0"),new TypedValue((int)DxfCode.Start,"LINE" ),new TypedValue((int)DxfCode.Operator, "and>"),new TypedValue((int)DxfCode.Operator,"<and"), new TypedValue((int)DxfCode.Start,"CIRCLE"), new TypedValue((int)DxfCode.Operator,">="),new TypedValue((int)DxfCode.Real,10.0), new TypedValue( (int)DxfCode.Operator,"and>"), new TypedValue((int)DxfCode.Operator, "or>") };


            acTypValAr.SetValue(new TypedValue((int)DxfCode.Operator, "<or"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 1); //lọc lấy polyline
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 2); //lọc lấy block
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Operator, "or>"), 3);
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult prSlRls = acEditor.GetSelection(acSelFtr);

            using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
            {
                if (prSlRls.Status == PromptStatus.OK)
                {
                    SelectionSet acSlSet = prSlRls.Value;
                    foreach (SelectedObject acSSObj in acSlSet)
                    {
                        Entity acEnt = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Entity;
                        ex = acEnt.GeometricExtents;

                        // Transform from UCS to DCS
                        ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1));
                        ResultBuffer rbTo = new ResultBuffer(new TypedValue(5003, 2));
                        double[] firres = new double[] { 0, 0, 0 };
                        double[] secres = new double[] { 0, 0, 0 };
                        // Transform the first point...
                        AcedTrans(ex.MinPoint.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
                        // ... and the second
                        AcedTrans(ex.MaxPoint.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
                        Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);
                        listWindows.Add(window);
                    }
                }
            }
            return listWindows;
        }
        public static List<Extents2d> GetWindow()   // tra ve mot list window
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDB = acDoc.Database;
            Editor acEditor = acDoc.Editor;
            List<Extents2d> listWindows = new List<Extents2d>();
            Point3d p1 = new Point3d(), p2 = new Point3d();

            // Create a TypedValue array to define the filter criteria
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0); //lọc lấy polyline
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 0); //lọc lấy block
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult prSlRls = acEditor.GetSelection(acSelFtr);

            using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
            {
                if (prSlRls.Status == PromptStatus.OK)
                {
                    SelectionSet acSlSet = prSlRls.Value;
                    foreach (SelectedObject acSSObj in acSlSet)
                    {
                        Polyline acEnt = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;
                        p1 = acEnt.GetPoint3dAt(0); //Min point
                        p2 = acEnt.GetPoint3dAt(2); //Max point

                        // Transform from UCS to DCS
                        ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1));
                        ResultBuffer rbTo = new ResultBuffer(new TypedValue(5003, 2));
                        double[] firres = new double[] { 0, 0, 0 };
                        double[] secres = new double[] { 0, 0, 0 };
                        // Transform the first point...
                        AcedTrans(p1.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
                        // ... and the second
                        AcedTrans(p2.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
                        Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);
                        listWindows.Add(window);
                    }
                }
            }
            return listWindows;
        }
        #endregion
        #region "Các hàm liên quan in ấn"
        public static StringCollection GetNameDevicePlot()   //tra ve danh sach ten may in
        {
            // gia tri mac dinh
            PlotSettingsValidator psv = PlotSettingsValidator.Current;
            StringCollection devList = psv.GetPlotDeviceList(); // lay danh sach may in
            return devList;
        }
        public static StringCollection GetPaperSize(string devName) //tra ve ten cac kho giay cua may in
        {
            StringCollection medList = new StringCollection();
            using (PlotSettings ps = new PlotSettings(true))
            {
                PlotSettingsValidator psv = PlotSettingsValidator.Current;
                psv.SetPlotConfigurationName(ps, devName, null);        // devName khong duoc gia tri null
                psv.RefreshLists(ps);
                medList = psv.GetCanonicalMediaNameList(ps);
            }
            return medList;
        }
        public static int CheckPlotToFileCap(string dev)
        //1: khong the; 2: bat buoc ; 3: co the
        //Kiểm tra khả năng xuất ra file của máy in.
        {
            PlotConfigManager.SetCurrentConfig(dev);
            PlotConfig plConf = PlotConfigManager.CurrentConfig;
            int ketqua = 0;
            if (plConf.PlotToFileCapability == PlotToFileCapability.NoPlotToFile)
            {
                ketqua = 1; //khong the in ra file
            }
            if (plConf.PlotToFileCapability == PlotToFileCapability.MustPlotToFile)
            {
                ketqua = 2; // bat buoc in ra file
            }
            if (plConf.PlotToFileCapability == PlotToFileCapability.PlotToFileAllowed)
            {
                ketqua = 3; // co the in ra file
            }
            return ketqua;
        }
        public static StringCollection GetNameStyle()//tra ve cac style in
        {
            PlotSettingsValidator psv = PlotSettingsValidator.Current;
            return psv.GetPlotStyleSheetList();
        }
        public static bool HuongKhoGiay(double x1, double y1, double x2, double y2) //true nam ngang, fail nam doc
        {
            if (Math.Abs(x1 - x2) > Math.Abs(y1 - y2)) return true;
            else return false;
        }
        public static void SimplePerivew(PlotInfo pi)  // in mot window
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            using (DocumentLock acLockDoc = acDoc.LockDocument())
            {

                using (PlotEngine pe = PlotFactory.CreatePreviewEngine((int)PreviewEngineFlags.Plot))
                {
                    //Validate the plot info
                    PlotInfoValidator piv = new PlotInfoValidator();
                    piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                    piv.Validate(pi);
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                    // Create a Progress Dialog to provide info and allow the user to cancel
                    PlotProgressDialog ppd = new PlotProgressDialog(true, 1, true);
                    using (ppd)
                    {
                        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Preview Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetName, acDoc.Name.Substring(acDoc.Name.LastIndexOf("\\") + 1));
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
                        ppd.LowerPlotProgressRange = 0;
                        ppd.UpperPlotProgressRange = 100;
                        ppd.PlotProgressPos = 0;

                        // Let's start the plot/preview, at last
                        ppd.OnBeginPlot();
                        ppd.IsVisible = true;
                        pe.BeginPlot(ppd, null);

                        //// We'll be plotting/previewing a single document
                        pe.BeginDocument(pi, acDoc.Name, null, 1, false, null);// Which contains a single sheet
                        ppd.OnBeginSheet();
                        ppd.LowerSheetProgressRange = 0;
                        ppd.UpperSheetProgressRange = 100;
                        ppd.SheetProgressPos = 0;

                        PlotPageInfo ppi = new PlotPageInfo();
                        pe.BeginPage(ppi, pi, true, null);
                        pe.BeginGenerateGraphics(null);
                        ppd.SheetProgressPos = 50;
                        pe.EndGenerateGraphics(null);

                        // Finish the sheet
                        PreviewEndPlotInfo pepi = new PreviewEndPlotInfo();
                        pe.EndPage(pepi);
                        ppd.SheetProgressPos = 100;
                        ppd.OnEndSheet();

                        // Finish the document
                        pe.EndDocument(null);

                        // And finish the plot
                        ppd.PlotProgressPos = 100;
                        ppd.OnEndPlot();
                        pe.EndPlot(null);
                    }
                }

            }
        }
        public static PlotInfo GetPlotInfo(string namePloter, string nameMedia, string nameStyle, Extents2d window)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor acEditor = acDoc.Editor;
            Database acCurDB = acDoc.Database;
            using (DocumentLock acLockDoc = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                {
                    LayoutManager acLayoutMgr = LayoutManager.Current;
                    DBDictionary myDic = acCurDB.LayoutDictionaryId.GetObject(OpenMode.ForRead) as DBDictionary;
                    String currentLayoutName = LayoutManager.Current.CurrentLayout as string;
                    Layout acLayout = new Layout();

                    if (myDic.Contains(currentLayoutName))
                    {
                        ObjectId layoutID = myDic.GetAt(currentLayoutName);
                        acLayout = layoutID.GetObject(OpenMode.ForRead) as Layout;
                    }

                    PlotSettings ps = new PlotSettings(acLayout.ModelType);
                    ps.CopyFrom(acLayout);

                    //update plot setting object
                    PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                    acPlSetVdr.SetPlotWindowArea(ps, window);
                    acPlSetVdr.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window); //chon vung in bang window
                    acPlSetVdr.SetCurrentStyleSheet(ps, nameStyle);
                    acPlSetVdr.SetUseStandardScale(ps, true);
                    acPlSetVdr.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                    acPlSetVdr.SetPlotCentered(ps, true);
                    acPlSetVdr.SetPlotConfigurationName(ps, namePloter, nameMedia);
                    if (HuongKhoGiay(window.MinPoint.X, window.MinPoint.Y, window.MaxPoint.X, window.MaxPoint.Y))
                        acPlSetVdr.SetPlotRotation(ps, PlotRotation.Degrees090);
                    else acPlSetVdr.SetPlotRotation(ps, PlotRotation.Degrees180);

                    //Set the plot info as an override since it will
                    //not be saved back to the layout
                    PlotInfo acPlInfo = new PlotInfo();
                    acPlInfo.Layout = acLayout.ObjectId;
                    acPlInfo.OverrideSettings = ps;
                    return acPlInfo;
                }
            }
        }
        public static void SimplePlot(PlotInfo pi, bool plotToFile, string pathFilePDF) //In một bản vẽ đơn
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor acEditor = acDoc.Editor;
            Database acCurDB = acDoc.Database;
            short bgPlot = (short)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("BACKGROUNDPLOT");
            using (DocumentLock acLockDoc = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                {
                    //update plot setting object
                    PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                    //Validate the plot info
                    PlotInfoValidator acPlInfoVdr = new PlotInfoValidator();
                    acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                    acPlInfoVdr.Validate(pi);
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                    //Check to see if a plot is already in progress
                    if (PlotFactory.ProcessPlotState == Autodesk.AutoCAD.PlottingServices.ProcessPlotState.NotPlotting)
                    {
                        using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())

                        //Track the plot progress with a Progress dialog
                        using (PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false, 1, true))
                        {
                            // Define the status messages to display when plotting starts
                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Plot progress");
                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");

                            //Set the plot progress  range
                            acPlProgDlg.LowerPlotProgressRange = 0;
                            acPlProgDlg.UpperPlotProgressRange = 10;
                            acPlProgDlg.PlotProgressPos = 0;

                            //Display the Progress dialog
                            acPlProgDlg.OnBeginPlot();
                            acPlProgDlg.IsVisible = true;

                            //Start to plot the layout
                            acPlEng.BeginPlot(acPlProgDlg, null);

                            // Define the plot output                            
                            if (plotToFile)
                            {
                                //acPlSet.PlotSettingsName = "fil1.pdf";
                                if (pathFilePDF == "" || pathFilePDF == null)
                                {
                                    pathFilePDF = Path.Combine(Path.GetDirectoryName(acCurDB.Filename), Path.GetFileNameWithoutExtension(acCurDB.Filename)) + ".pdf";
                                }
                                string outFile = pathFilePDF;
                                acPlEng.BeginDocument(pi, acDoc.Name, null, 1, true, outFile);
                            }
                            else acPlEng.BeginDocument(pi, acDoc.Name, null, 1, false, null);

                            // Display information about the current plot
                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.Status, "Plotting:" + acDoc.Name + "-" + pi.Layout.ToString());

                            // Set the sheet progress range
                            acPlProgDlg.OnBeginSheet();
                            acPlProgDlg.LowerSheetProgressRange = 0;
                            acPlProgDlg.UpperSheetProgressRange = 100;
                            acPlProgDlg.SheetProgressPos = 0;

                            //Plot the first sheet/layout
                            PlotPageInfo acPlPageInfo = new PlotPageInfo();
                            acPlEng.BeginPage(acPlPageInfo, pi, true, null);

                            acPlEng.BeginGenerateGraphics(null);
                            acPlEng.EndGenerateGraphics(null);

                            // Finish plotting the sheet/layout
                            acPlEng.EndPage(null);
                            acPlProgDlg.SheetProgressPos = 100;
                            acPlProgDlg.OnEndSheet();

                            //Finish plotting the document
                            acPlEng.EndDocument(null);
                            // Finish the plot
                            acPlProgDlg.PlotProgressPos = 100;
                            acPlProgDlg.OnEndPlot();
                            acPlEng.EndPlot(null);
                        }
                    }
                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", bgPlot);
                }
            }
        }
        #endregion
        #region "Các hàm vẽ"
        
        private static void AddToModelSpace(Entity ent, Database db)    //Them doi tuong vao database cua ban ve
        {

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                modelSpace.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent, true);
                trans.Commit();
            }
        }
        public static void InserText(double startPointX, double startPointY, double height, string content) // chen text vao ban ve
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            using (DocumentLock acLockDoc = doc.LockDocument())
            {
                //align point
                Point3d pt = new Point3d(startPointX, startPointY, 0); ;

                using (DBText text = new DBText())
                {
                    text.TextString = content;
                    text.Height = height;
                    text.VerticalMode = TextVerticalMode.TextVerticalMid;
                    text.HorizontalMode = TextHorizontalMode.TextCenter;
                    text.AlignmentPoint = pt;
                    AddToModelSpace(text, db);
                }
            }
        }
        public static void InserMText(double startPointX, double startPointY, double height, string content)//Chen Multi line Text
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (acDoc != null)
            {

                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        //Open the Block table for read
                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                        //Open the Block table record Model space for write
                        BlockTableRecord acBlkTblRec;
                        acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        //Create a multiline text object
                        using (MText acMText = new MText())
                        {
                            acMText.Location = new Point3d(startPointX, startPointY, 0);
                            acMText.TextHeight = height;
                            acMText.Attachment = AttachmentPoint.MiddleCenter;
                            acMText.Contents = content;
                            acBlkTblRec.AppendEntity(acMText);
                            acTrans.AddNewlyCreatedDBObject(acMText, true);
                        }

                        //Save the changes and dispose of the transaction
                        acTrans.Commit();
                    }
                }

            }
            //Start a transaction

        }
        public static void AddLine(Line l1) //Vẽ line vào bản vẽ hiện hành
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        btr.AppendEntity(l1);
                        acTrans.AddNewlyCreatedDBObject(l1, true);
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void AddLine(Line l1,Document acDoc)
        {            
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        btr.AppendEntity(l1);
                        acTrans.AddNewlyCreatedDBObject(l1, true);
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void AddLine(Line l1, Document acDoc,LayerTableRecord layer)
        {
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
               
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        
                        BlockTable bt = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        l1.Layer = layer.Name;
                        btr.AppendEntity(l1);
                        acTrans.AddNewlyCreatedDBObject(l1, true);
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void AddLine(Line l1, Document acDoc, string layerName)
        {
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable layerTable = (LayerTable)acTrans.GetObject(acCurDB.LayerTableId, OpenMode.ForRead);
                        if(layerTable.Has(layerName))
                        {
                            l1.Layer = layerName;
                        }                        
                        btr.AppendEntity(l1);
                        acTrans.AddNewlyCreatedDBObject(l1, true);
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void AddRec2p(double x1, double y1, double x2, double y2) // them rectange biet diem dau diem cuoi
        {
            Point2d p1 = new Point2d(x1, y1);
            Point2d p2 = new Point2d(x2, y1);
            Point2d p3 = new Point2d(x1, y2);
            Point2d p4 = new Point2d(x2, y2);
            var pl1 = new Polyline(4);
            pl1.AddVertexAt(0, p1, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(1, p2, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(2, p4, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(3, p3, 0.0, -1.0, -1.0);
            pl1.Closed = true;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        btr.AppendEntity(pl1);
                        acTrans.AddNewlyCreatedDBObject(pl1, true);
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void AddRec2p(double x1, double y1, double x2, double y2, string layerName) // them rectange biet diem dau diem cuoi
        {
            Point2d p1 = new Point2d(x1, y1);
            Point2d p2 = new Point2d(x2, y1);
            Point2d p3 = new Point2d(x1, y2);
            Point2d p4 = new Point2d(x2, y2);
            var pl1 = new Polyline(4);
            pl1.Layer = layerName;
            pl1.AddVertexAt(0, p1, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(1, p2, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(2, p4, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(3, p3, 0.0, -1.0, -1.0);
            pl1.Closed = true;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        btr.AppendEntity(pl1);
                        acTrans.AddNewlyCreatedDBObject(pl1, true);
                        acTrans.Commit();
                    }
                }

            }
        }
        public static void AddRec2p(double x1, double y1, double x2, double y2, Color color) // them rectange biet diem dau diem cuoi
        {
            Point2d p1 = new Point2d(x1, y1);
            Point2d p2 = new Point2d(x2, y1);
            Point2d p3 = new Point2d(x1, y2);
            Point2d p4 = new Point2d(x2, y2);
            var pl1 = new Polyline(4);
            pl1.Color = color;
            pl1.AddVertexAt(0, p1, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(1, p2, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(2, p4, 0.0, -1.0, -1.0);
            pl1.AddVertexAt(3, p3, 0.0, -1.0, -1.0);
            pl1.Closed = true;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDB = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDB.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = acTrans.GetObject(acCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = acTrans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        btr.AppendEntity(pl1);
                        acTrans.AddNewlyCreatedDBObject(pl1, true);
                        acTrans.Commit();
                    }
                }

            }
        }
        #endregion
        #region "Các Style"
        public static void LoadLineType(string sLineTypName) //Hàm chọn nét vẽ cho bản vẽ hiện hành.
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            using (DocumentLock acLockDoc = doc.LockDocument())
            {
                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    // Open the Linetype table for read
                    LinetypeTable acLineTypTbl;
                    acLineTypTbl = acTrans.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                    if (acLineTypTbl.Has(sLineTypName) == false)
                    {
                        // Load the Center Linetype
                        db.LoadLineTypeFile(sLineTypName, "acad.lin");
                    }
                    // Save the changes and dispose of the transaction
                    acTrans.Commit();
                }
            }
        }
        public static void CreatLayer(string layerName, string sLineTypName, short colorIndex, LineWeight lineWeight, bool isPlottAble) //Creat a layer
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            using (DocumentLock acLockDoc = doc.LockDocument())
            {
                using (Transaction acTrans = db.TransactionManager.StartTransaction())
                {
                    // Get the layer table from the drawing
                    // Check the layer name, to see whether it's
                    LayerTable lt = (LayerTable)acTrans.GetObject(db.LayerTableId, OpenMode.ForRead);
                    LinetypeTable acLinTbl = acTrans.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                    SymbolUtilityServices.ValidateSymbolName(layerName, false);
                    if (!lt.Has(layerName))
                    {
                        // Create our new layer table record...
                        // ... and set its properties
                        LayerTableRecord ltr = new LayerTableRecord();
                        ltr.Name = layerName;
                        ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                        ltr.LineWeight = lineWeight;
                        ltr.IsPlottable = isPlottAble;
                        ltr.LinetypeObjectId = acLinTbl[sLineTypName];
                        // Add the new layer to the layer table
                        lt.UpgradeOpen();
                        ObjectId ltId = lt.Add(ltr);
                        acTrans.AddNewlyCreatedDBObject(ltr, true);
                    }
                    //else  MessageBox.Show( "\nA layer with this name already exists.");
                    acTrans.Commit();
                }
            }
        }
        public static void CreatTextStyle(string nameTextStyle, string nameFont, double textSize) //Creat a Text Style
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDb = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        TextStyleTable newTextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                        if (!newTextStyleTable.Has(nameTextStyle))
                        {
                            newTextStyleTable.UpgradeOpen();
                            TextStyleTableRecord newTextStyleTableRecord = new TextStyleTableRecord();
                            newTextStyleTableRecord.FileName = nameFont;

                            Autodesk.AutoCAD.GraphicsInterface.FontDescriptor oldFont = newTextStyleTableRecord.Font;
                            Autodesk.AutoCAD.GraphicsInterface.FontDescriptor newFont = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(oldFont.TypeFace, false, false, oldFont.CharacterSet, oldFont.PitchAndFamily);
                            newTextStyleTableRecord.Font = newFont;
                            newTextStyleTableRecord.TextSize = textSize;
                            newTextStyleTableRecord.Name = nameTextStyle;
                            newTextStyleTable.Add(newTextStyleTableRecord);
                            acTrans.AddNewlyCreatedDBObject(newTextStyleTableRecord, true);
                        }
                        // Save the new object to the database
                        acTrans.Commit();
                    }
                }
            }
        }
        public static TextStyleTable GetTextStyleTable() //Hàm trả về bảng các text Style
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            TextStyleTable newTextStyleTable = null;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDb = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        newTextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    }
                }
            }
            return newTextStyleTable;
        }
        public static void CreatDimStyle(string nameDimStyle, string textStyle) 
            //Text Heigh = dimScale x 2.5
            //Tao Dim Style tieu chuan
        {
            TextStyleTable acTextStyleTable = GetTextStyleTable();
            if (!acTextStyleTable.Has(textStyle))
            {
                MessageBox.Show("Chưa khởi tạo:" + textStyle);
                return;
            }
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDb = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        DimStyleTable newDimStyleTable = acTrans.GetObject(acCurDb.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                        TextStyleTableRecord txtStyleRecord = acTrans.GetObject(acTextStyleTable[textStyle], OpenMode.ForRead) as TextStyleTableRecord;
                        if (!newDimStyleTable.Has(nameDimStyle))
                        {
                            newDimStyleTable.UpgradeOpen();
                            DimStyleTableRecord newDimStyleTableRecord = new DimStyleTableRecord();
                            newDimStyleTableRecord.Name = nameDimStyle;
                            // *** LINES tab ***
                            // "Dimension lines" group:
                            newDimStyleTableRecord.Dimclrd = Color.FromColorIndex(ColorMethod.ByAci, 2); // Color cho đường base
                            newDimStyleTableRecord.Dimltype = acCurDb.ByBlockLinetype;                  // Linetype
                            newDimStyleTableRecord.Dimlwd = LineWeight.ByBlock; // Lineweight;
                            newDimStyleTableRecord.Dimdle = 0;                  // Extend Beyond Ticks
                            newDimStyleTableRecord.Dimdli = txtStyleRecord.TextSize * 2;                  // Baseline Spacing
                            newDimStyleTableRecord.Dimsd1 = false;              // Suppress dim line 1
                            newDimStyleTableRecord.Dimsd2 = false;                            // Suppress dim line 2

                            // "Extension Lines" group:(Đường gióng kích thước)
                            newDimStyleTableRecord.Dimclre = Color.FromColorIndex(ColorMethod.ByAci, 3); // Color cho đường gióng
                            newDimStyleTableRecord.Dimltex1 = acCurDb.ByBlockLinetype;  // Linetype Ext 1
                            newDimStyleTableRecord.Dimltex2 = acCurDb.ByBlockLinetype;  // Linetype Ext 2
                            newDimStyleTableRecord.Dimlwe = LineWeight.LineWeight013;         // Lineweight
                            newDimStyleTableRecord.Dimse1 = false;                      // Suppress Ext line 1
                            newDimStyleTableRecord.Dimse2 = false;                      // Suppress Ext line 1
                            newDimStyleTableRecord.Dimexe = txtStyleRecord.TextSize / 3;        // Extend Beyond Dim Lines
                            newDimStyleTableRecord.Dimexo = txtStyleRecord.TextSize / 3;       // Offset From Origin
                            newDimStyleTableRecord.DimfxlenOn = false; // Fixed Length Extension Lines
                            newDimStyleTableRecord.Dimfxlen = 1; // Length

                            // *** SYMBOL AND ARROWS tab ***
                            // "Arrowheads" group:
                            // Note: Annotative blocks cannot be used
                            // as custom arrowheads for dimensions or 
                            // leaders.
                            // Arrows" tab
                            //newDimStyleTableRecord.Dimblk1 = id1;
                            newDimStyleTableRecord.Dimasz = txtStyleRecord.TextSize; // Arrow Size

                            // "Center marks" group:  
                            // 'Dimcen' property's allowed values:
                            //  0 - None; 
                            //  1 - Mark; 
                            // -1 - Line
                            Int32 centerMarks = -1;
                            Double centerMarksSize = 500;
                            newDimStyleTableRecord.Dimcen = centerMarks * centerMarksSize;

                            // *** TEXT tab ***
                            // "Text Appearance" group: 
                            // Text Style
                            newDimStyleTableRecord.Dimtxsty = txtStyleRecord.ObjectId; //Text Style
                            newDimStyleTableRecord.Dimclrt = Color.FromColorIndex(ColorMethod.ByAci, 2); // Color cho text
                            // 'Dimtfill' property's allowed values:
                            // 0 - No background
                            // 1 - The background color of the 
                            //    drawing
                            // 2 - The background specified by 
                            //    Dimtfillclr property
                            newDimStyleTableRecord.Dimtfill = 0;
                            //newDimStyleTableRecord.Dimtfillclr = Color.FromColorIndex(ColorMethod.ByAci, 2);

                            //newDimStyleTableRecord.Dimtxt = 3.5; // Text Height
                            //newDimStyleTableRecord.Dimfrac = 2;

                            // "Text Placement" group:
                            // 'Dimtad' property's allowed values:
                            // 0 - Centers the dimension text between the extension lines. ( chính giữa )
                            // 1 - Places the dimension text above  the dimension line except when the 
                            // dimension line is not horizontal and text inside the extension lines
                            // is forced horizontal (DIMTIH = 1). The distance from the dimension line to the baseline of the lowest 
                            //    line of text is the current DIMGAP value.
                            // 2 - Places the dimension text on the 
                            //    side of the dimension line farthest 
                            //    away from the defining points.
                            // 3 - Places the dimension text to 
                            //    conform to Japanese Industrial 
                            //    Standards (JIS).
                            newDimStyleTableRecord.Dimtad = 1; // Vertical 
                            newDimStyleTableRecord.Dimjust = 0; // Horizontal 0=center
                            newDimStyleTableRecord.Dimgap = txtStyleRecord.TextSize / 3;// Offset from Dim Line

                            // "Text Alignment" group:
                            newDimStyleTableRecord.Dimtih = false; // Text Alignment
                            newDimStyleTableRecord.Dimtoh = true; //( hai giá trị này khác nhau thì sẽ chọn kiểu ISO

                            // *** FIT tab ***
                            /* 'Dimjust' property's allowed values:    
                             0 - Places both text and arrows outside extension lines
                             1 - Moves arrows first, then text
                             2 - Moves text first, then arrows
                             3 - Moves either text or arrows,whichever fits best */
                            // The 'Dimtix' property must be false
                            newDimStyleTableRecord.Dimatfit = 3;
                            newDimStyleTableRecord.Dimtix = false;        //Always Keep Text Between Ext Lines
                            newDimStyleTableRecord.Dimsoxd = false;       // Suppress Arrows If They Don't Fit Inside Extension Lines

                            // "Text placement" group:
                            // 0 - Moves the dimension line with dimension text
                            // 1 - Adds a leader when dimension text is moved
                            // 2 - Allows text to be moved freely without a leader 
                            newDimStyleTableRecord.Dimtmove = 0;
                            newDimStyleTableRecord.Dimtofl = true; // Draw Dim Line Between Ext Lines

                            // *** PRIMARY UNITS tab ***
                            // 'Dimlunit' property's allowed values:
                            // 1 - Scientific
                            // 2 - Decimal
                            // 3 - Engineering
                            // 4 - Architectural (always displayed stacked)
                            // 5 - Fractional (always displayed stacked)
                            // 6 - Microsoft Windows Desktop (decimal format using Control Panel

                            // Linear Dimensions
                            newDimStyleTableRecord.Dimlunit = 2;
                            newDimStyleTableRecord.Dimdec = 2; // Precision
                            // The relative sizes of numbers in stacked fractions
                            // are based on the DIMTFAC system variable (in the same 
                            // way that tolerance values use this system variable).
                            newDimStyleTableRecord.Dimtfac = 1.0;
                            newDimStyleTableRecord.Dimdsep = '.';
                            newDimStyleTableRecord.Dimrnd = 0; // Round Off
                            newDimStyleTableRecord.Dimzin = 8;
                            newDimStyleTable.Add(newDimStyleTableRecord);
                            acTrans.AddNewlyCreatedDBObject(newDimStyleTableRecord, true);
                        }
                        // Save the new object to the database
                        acTrans.Commit();
                    }
                }
            }
        }
        #endregion
        #region "Các hàm liên quan đến dữ liệu"
        public static void CreatCostumProperties(string key, string value) //Tạo costum properties
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDb = acDoc.Database;

                DatabaseSummaryInfo info;
                DatabaseSummaryInfoBuilder builder = new DatabaseSummaryInfoBuilder();

                //Copy custom properties cu
                IDictionaryEnumerator custProp = acCurDb.SummaryInfo.CustomProperties;
                while (custProp.MoveNext())
                {
                    builder.CustomPropertyTable.Add(custProp.Key, custProp.Value);
                }
                //Them key moi vao
                if (!builder.CustomPropertyTable.Contains(key))
                {
                    builder.CustomPropertyTable.Add(key, value);
                    info = builder.ToDatabaseSummaryInfo();
                    acCurDb.SummaryInfo = info;
                }
            }
        }
        #endregion
    }
}
