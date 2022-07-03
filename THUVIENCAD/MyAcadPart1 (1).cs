using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk;
using System.Windows.Forms;
using Autodesk.AutoCAD.Colors;
using System.Collections;

namespace ACADTRANSFORMER.THUVIENCAD
{
    partial class MyAcad
    {
        public static void DrawLaGong(Point2d p1, double khoton, double chieudai)
        {
            Point2d p2 = new Point2d(p1.X + chieudai + khoton, p1.Y);
            Point2d p3 = new Point2d(p2.X - khoton, p2.Y + khoton);
            Point2d p4 = new Point2d(p3.X - (chieudai / 2 - khoton), p3.Y);
            Point2d p5 = new Point2d(p4.X - (khoton / 2), p4.Y - khoton / 2);
            Point2d p6 = new Point2d(p5.X - (khoton / 2), p5.Y + khoton / 2);
            Point2d p7 = new Point2d(p6.X - (chieudai / 2 - khoton), p6.Y);
            var poliline1 = new Polyline(7);
            poliline1.AddVertexAt(0, p1, 0.0, -1.0, -1.0);
            poliline1.AddVertexAt(1, p2, 0.0, -1.0, -1.0);
            poliline1.AddVertexAt(2, p3, 0.0, -1.0, -1.0);
            poliline1.AddVertexAt(3, p4, 0.0, -1.0, -1.0);
            poliline1.AddVertexAt(4, p5, 0.0, -1.0, -1.0);
            poliline1.AddVertexAt(5, p6, 0.0, -1.0, -1.0);
            poliline1.AddVertexAt(6, p7, 0.0, -1.0, -1.0);
            poliline1.Closed = true;

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
                        //Lấy bảng layer
                        LayerTable acLyrTbl = acTrans.GetObject(acCurDB.LayerTableId, OpenMode.ForRead) as LayerTable;
                        //Lay cac dimenison style
                        DimStyleTable acDimstyleTbl = acTrans.GetObject(acCurDB.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                        //Ve duong tam 1
                        Point3d diem1 = new Point3d(p5.X + chieudai / 2 + 60, p5.Y, 0);
                        Point3d diem2 = new Point3d(p5.X - chieudai / 2 - 60, p5.Y, 0);
                        Line l1 = new Line(diem1, diem2);
                        l1.Layer = "Centerline (ISO)";
                        l1.LinetypeScale = 66;
                        btr.AppendEntity(l1);
                        acTrans.AddNewlyCreatedDBObject(l1, true);

                        //Ve duong tam 2
                        Point3d diem3 = new Point3d(diem1.X, p5.Y + 25, 0);
                        Point3d diem4 = new Point3d(diem2.X, p5.Y + 25, 0);
                        Line l2 = new Line(diem3, diem4);
                        l2.Layer = "Centerline (ISO)";
                        l2.LinetypeScale = 66;
                        btr.AppendEntity(l2);
                        acTrans.AddNewlyCreatedDBObject(l2, true);

                        //Vẽ các đường kích thước
                        using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear C
                        {
                            acRotDim.XLine1Point = diem1;
                            acRotDim.XLine2Point = diem3;
                            acRotDim.Rotation = Math.PI / 2;
                            acRotDim.DimLinePoint = new Point3d(diem1.X+80, diem1.Y + 180, 0);
                            acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                            acRotDim.Layer = "Dimension (ISO)";
                            acRotDim.DimensionText = "C";
                            // Add the new object to Model space and the transaction
                            btr.AppendEntity(acRotDim);
                            acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                        }
                        using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear A
                        {
                            acRotDim.XLine1Point = new Point3d(p1.X, p1.Y, 0);
                            acRotDim.XLine2Point = new Point3d(p7.X, p7.Y, 0);
                            acRotDim.Rotation = Math.PI / 2;
                            acRotDim.DimLinePoint = new Point3d(p1.X - 50, (p1.Y + p7.Y) / 2, 0);
                            acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                            acRotDim.Layer = "Dimension (ISO)";
                            acRotDim.DimensionText = "A";
                            // Add the new object to Model space and the transaction
                            btr.AppendEntity(acRotDim);
                            acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                        }
                        using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear L2
                        {
                            acRotDim.XLine1Point = new Point3d((p1.X + p7.X) / 2, (p1.Y + p7.Y) / 2, 0);
                            acRotDim.XLine2Point = new Point3d(p5.X, p5.Y, 0);
                            acRotDim.Rotation = 0;
                            acRotDim.DimLinePoint = new Point3d(p5.X - chieudai / 4, p5.Y - 230, 0);
                            acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                            acRotDim.DimensionText = "L2";
                            acRotDim.Layer = "Dimension (ISO)";
                            // Add the new object to Model space and the transaction
                            btr.AppendEntity(acRotDim);
                            acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                        }
                        using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear L2'
                        {
                            acRotDim.XLine1Point = new Point3d(p5.X + chieudai / 2, p5.Y, 0);
                            acRotDim.XLine2Point = new Point3d(p5.X, p5.Y, 0);
                            acRotDim.Rotation = 0;
                            acRotDim.DimLinePoint = new Point3d(p5.X + chieudai / 4, p5.Y - 230, 0);
                            acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                            acRotDim.DimensionText = "L2";
                            acRotDim.Layer = "Dimension (ISO)";
                            // Add the new object to Model space and the transaction
                            btr.AppendEntity(acRotDim);
                            acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                        }
                        using (LineAngularDimension2 acAngular = new LineAngularDimension2()) // Ve kich thuoc goc 90 do
                        {
                            acAngular.XLine1Start = new Point3d(p5.X, p5.Y, 0);
                            acAngular.XLine1End = new Point3d(p6.X, p6.Y, 0);
                            acAngular.XLine2Start = new Point3d(p5.X, p5.Y, 0);
                            acAngular.XLine2End = new Point3d(p4.X, p4.Y, 0);
                            acAngular.ArcPoint = new Point3d(p5.X, p5.Y + 200, 0);
                            acAngular.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                            acAngular.Layer = "Dimension (ISO)";
                            // Add the new object to Model space and the transaction
                            btr.AppendEntity(acAngular);
                            acTrans.AddNewlyCreatedDBObject(acAngular, true);
                        }


                        poliline1.Layer = "Visible (ISO)";

                        btr.AppendEntity(poliline1);
                        acTrans.AddNewlyCreatedDBObject(poliline1, true);

                        acTrans.Commit();
                    }
                }
            }
        }
        public static void DrawTruBen(Point2d p1 ,double khoton,double chieudai,bool isVisibleDim)
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
                        //Lấy bảng layer
                        LayerTable acLyrTbl = acTrans.GetObject(acCurDB.LayerTableId, OpenMode.ForRead) as LayerTable;
                        //Lay cac dimenison style
                        DimStyleTable acDimstyleTbl = acTrans.GetObject(acCurDB.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;

                        Point2d p2 = new Point2d(p1.X + chieudai - khoton, p1.Y);
                        Point2d p3 = new Point2d(p2.X + khoton, p2.Y + khoton);
                        Point2d p4 = new Point2d(p3.X - chieudai - khoton, p3.Y);
                        var poliline1 = new Polyline(4);
                        poliline1.AddVertexAt(0, p1, 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(1, p2, 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(2, p3, 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(3, p4, 0.0, -1.0, -1.0);
                        poliline1.Closed = true;
                        poliline1.Layer = "Visible (ISO)";
                        btr.AppendEntity(poliline1);
                        acTrans.AddNewlyCreatedDBObject(poliline1, true);

                        //Ve duong tam
                        Line l1 = new Line(new Point3d((p1.X+p4.X)/2-60,(p1.Y+p4.Y)/2,0),new Point3d((p2.X+p3.X)/2+60,(p2.Y+p3.Y)/2,0));
                        l1.Layer = "Centerline (ISO)";
                        l1.LinetypeScale = 66;
                        btr.AppendEntity(l1);
                        acTrans.AddNewlyCreatedDBObject(l1, true);

                        if (isVisibleDim)
                        {
                            using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear B
                            {
                                acRotDim.XLine1Point = new Point3d(p2.X, p2.Y, 0);
                                acRotDim.XLine2Point = new Point3d(p3.X, p3.Y, 0);
                                acRotDim.Rotation = Math.PI / 2;
                                acRotDim.DimLinePoint = new Point3d(p3.X + 50, (p3.Y + p2.Y) / 2, 0);
                                acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                                acRotDim.Layer = "Dimension (ISO)";
                                acRotDim.DimensionText = "B";
                                // Add the new object to Model space and the transaction
                                btr.AppendEntity(acRotDim);
                                acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                            }
                            using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear L1
                            {
                                acRotDim.XLine1Point = new Point3d((p1.X+p4.X)/2, (p1.Y+p4.Y)/2, 0);
                                acRotDim.XLine2Point = new Point3d((p2.X + p3.X) / 2, (p2.Y + p3.Y) / 2, 0);
                                acRotDim.Rotation = 0.0;
                                acRotDim.DimLinePoint = new Point3d((p1.X + p2.X)/2, p1.Y-137.5, 0);
                                acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                                acRotDim.Layer = "Dimension (ISO)";
                                acRotDim.DimensionText = "L1";
                                // Add the new object to Model space and the transaction
                                btr.AppendEntity(acRotDim);
                                acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                            }
                            using (LineAngularDimension2 acAngular = new LineAngularDimension2()) // Ve kich thuoc goc 45 do
                            {
                                acAngular.XLine1Start = new Point3d(p3.X, p3.Y, 0);
                                acAngular.XLine2Start = new Point3d(p3.X, p3.Y, 0);
                                acAngular.XLine1End = new Point3d(p4.X, p4.Y, 0);                                
                                acAngular.XLine2End = new Point3d(p2.X, p2.Y, 0);
                                acAngular.ArcPoint = new Point3d(p2.X+15, p2.Y + 87, 0);
                                acAngular.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                                acAngular.Layer = "Dimension (ISO)";
                                // Add the new object to Model space and the transaction
                                btr.AppendEntity(acAngular);
                                acTrans.AddNewlyCreatedDBObject(acAngular, true);
                            }
                        }
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void DrawTruGiua(Point3d p1, bool isVisibleDim)
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
                        //Lấy bảng layer
                        LayerTable acLyrTbl = acTrans.GetObject(acCurDB.LayerTableId, OpenMode.ForRead) as LayerTable;
                        //Lay cac dimenison style
                        DimStyleTable acDimstyleTbl = acTrans.GetObject(acCurDB.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;

                        Point3d p2 = new Point3d(p1.X + 1200, p1.Y,0);
                        Point3d p3 = new Point3d(p1.X , p1.Y -30,0);
                        Point3d p4 = new Point3d(p2.X , p2.Y-30,0);
                        Point3d p5 = new Point3d(p1.X, p1.Y + 30,0);
                        Point3d p6 = new Point3d(p2.X, p2.Y + 30,0);

                        Point2d p7 = new Point2d(p5.X+85, p5.Y + 85);
                        Point2d p8 = new Point2d(p7.X+970, p7.Y);
                        Point2d p9 = new Point2d(p4.X-85, p4.Y-85);
                        Point2d p10 = new Point2d(p9.X-970, p9.Y);
                        
                        //Ve duong tam
                        Line center1 = new Line(p1,p2);
                        Line center2 = new Line(p3, p4);
                        Line center3 = new Line(p5, p6);
                        center1.Layer = "Centerline (ISO)";
                        center2.Layer = "Centerline (ISO)";
                        center3.Layer = "Centerline (ISO)";
                        center1.LinetypeScale = 66;
                        center2.LinetypeScale = 66;
                        center3.LinetypeScale = 66;
                        btr.AppendEntity(center1);
                        btr.AppendEntity(center2); 
                        btr.AppendEntity(center3);
                        acTrans.AddNewlyCreatedDBObject(center1, true);
                        acTrans.AddNewlyCreatedDBObject(center2, true);
                        acTrans.AddNewlyCreatedDBObject(center3, true);

                        var poliline1 = new Polyline(6);
                        poliline1.AddVertexAt(0,new Point2d(p5.X,p5.Y), 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(1, p7, 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(2, p8, 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(3, new Point2d(p4.X, p4.Y), 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(4, p9, 0.0, -1.0, -1.0);
                        poliline1.AddVertexAt(5, p10, 0.0, -1.0, -1.0);
                        poliline1.Closed = true;
                        poliline1.Layer = "Visible (ISO)";
                        btr.AppendEntity(poliline1);
                        acTrans.AddNewlyCreatedDBObject(poliline1, true);

                        if (isVisibleDim)
                        {
                            using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear L1
                            {
                                acRotDim.XLine1Point = p4;
                                acRotDim.XLine2Point = p5;
                                acRotDim.Rotation = 0.0;
                                acRotDim.DimLinePoint = new Point3d((p1.X + p2.X)/2, p1.Y -260, 0);
                                acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                                acRotDim.Layer = "Dimension (ISO)";
                                acRotDim.DimensionText = "L1";
                                // Add the new object to Model space and the transaction
                                btr.AppendEntity(acRotDim);
                                acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                            }
                            using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear B
                            {
                                acRotDim.XLine1Point = new Point3d(p7.X ,p7.Y, 0);
                                acRotDim.XLine2Point = new Point3d(p10.X, p10.Y, 0);
                                acRotDim.Rotation = Math.PI/2;
                                acRotDim.DimLinePoint = new Point3d(p1.X -100, p1.Y,0);
                                acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                                acRotDim.Layer = "Dimension (ISO)";
                                acRotDim.DimensionText = "B";
                                // Add the new object to Model space and the transaction
                                btr.AppendEntity(acRotDim);
                                acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                            }
                            using (RotatedDimension acRotDim = new RotatedDimension()) //Dimlinear D
                            {
                                acRotDim.XLine1Point = new Point3d(p2.X, p2.Y, 0);
                                acRotDim.XLine2Point = new Point3d(p4.X, p4.Y, 0);
                                acRotDim.Rotation = Math.PI / 2;
                                acRotDim.DimLinePoint = new Point3d(p2.X +90, p2.Y+160, 0);
                                acRotDim.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                                acRotDim.Layer = "Dimension (ISO)";
                                acRotDim.DimensionText = "D";
                                // Add the new object to Model space and the transaction
                                btr.AppendEntity(acRotDim);
                                acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                            }
                            using (LineAngularDimension2 acAngular = new LineAngularDimension2()) // Ve kich thuoc goc 45 do
                            {
                                acAngular.XLine1Start = new Point3d(p5.X, p5.Y, 0);
                                acAngular.XLine2Start = new Point3d(p5.X, p5.Y, 0);
                                acAngular.XLine1End = new Point3d(p10.X, p10.Y, 0);
                                acAngular.XLine2End = new Point3d(p7.X, p7.Y, 0);
                                acAngular.ArcPoint = new Point3d(p5.X + 300, p5.Y+30, 0);
                                acAngular.DimensionStyle = acDimstyleTbl["X26.4_DUONG"];
                                acAngular.Layer = "Dimension (ISO)";
                                // Add the new object to Model space and the transaction
                                btr.AppendEntity(acAngular);
                                acTrans.AddNewlyCreatedDBObject(acAngular, true);
                            }
                        }
                        acTrans.Commit();
                    }
                }
            }
        }
    }
}
