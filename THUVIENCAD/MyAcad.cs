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
        public static void InserMText(double startPointX, double startPointY, double height, string content)
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
        public static Polyline CreatRectangle(double x1, double y1, double x2, double y2)
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
            return pl1;
        }
        public static Line CreatLine(double startPointX, double startPointY, double endPointX, double endPointY)
        {
            Line myLine = new Line(new Point3d(startPointX, startPointY, 0), new Point3d(endPointX, endPointY, 0));
            return  myLine;
        }
        public static void AddLine(Line l1)
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
        public static Circle CreatCircle(double center_x, double center_y, double radial)
        {
            Point3d p = new Point3d(center_x, center_y, 0);
            Circle c1 = new Circle(p, Vector3d.ZAxis, radial);
            return c1;
        }
        public static void AddCircle(Point3d cPoint, double rad, Color myColor)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Circle c1 = new Circle(cPoint, Vector3d.ZAxis, rad);
            c1.Color = myColor;
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
                        btr.AppendEntity(c1);
                        acTrans.AddNewlyCreatedDBObject(c1, true);
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void AddCircle(Point3d cPoint, double rad)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Circle c1 = new Circle(cPoint, Vector3d.ZAxis, rad);
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
                        btr.AppendEntity(c1);
                        acTrans.AddNewlyCreatedDBObject(c1, true);
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void AddCircle(double center_x, double center_y, double radial)
        {
            Point3d p = new Point3d(center_x, center_y, 0);
            AddCircle(p, radial);
        }
        public static void AddCircle(double center_x, double center_y, double radial, Color colorIndex)
        {
            Point3d p = new Point3d(center_x, center_y, 0);
            AddCircle(p, radial, colorIndex);
        }
        public static void AddrecWidthHigh(double x, double y, double with, double high)//ve hinh vuong biet toa do goc duoi ben trai va chieu rong, chieu cao
        {
            AddRec2p(x, y, x + with, y + high);
        }
        public static void AddrecWidthHigh(double x, double y, double with, double high, Color color)//ve hinh vuong biet toa do goc duoi ben trai va chieu rong, chieu cao
        {
            AddRec2p(x, y, x + with, y + high, color);
        }
        public static double[] GetPosition()
        {
            double[] ketqua = new double[] { 0, 0, 0 }; //phan tu thu 3 tra ve trang thai cua prPoint
            Point3d p1 = new Point3d();
            PromptPointResult prPointRls;
            PromptPointOptions prPointOpts = new PromptPointOptions("Chọn tọa độ tâm của mặt cắt");

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDB = acDoc.Database;
            Editor acEditor = acDoc.Editor;

            using (DocumentLock acLockDoc = acDoc.LockDocument())
            {
                prPointRls = acEditor.GetPoint(prPointOpts);
                if (prPointRls.Status == PromptStatus.OK)
                {
                    p1 = prPointRls.Value;
                    ketqua[0] = p1.X;
                    ketqua[1] = p1.Y;
                    ketqua[2] = 1; //thanh cong
                }
                return ketqua;
            }
        }
        public static void AddLine2d(double startPointX, double startPointY, double endPointX, double endPointY)
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
                        Line myLine = new Line(new Point3d(startPointX, startPointY, 0), new Point3d(endPointX, endPointY, 0));
                        btr.AppendEntity(myLine);
                        acTrans.AddNewlyCreatedDBObject(myLine, true);
                        acTrans.Commit();
                    }
                }
            }
        }        
        public static void DefineBlockBorderA3(string nameBlock)
        {
            double startPointX=0, startPointY=0;

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDb = acDoc.Database;
                using (DocumentLock acLockDoc = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        // Open the Block table for read
                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        ObjectId blkRecId = ObjectId.Null;
                        
                        if (!acBlkTbl.Has(nameBlock))    // Neu chua co block border A3
                        {
                            using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                            {
                                acBlkTblRec.Name = nameBlock;

                                // Set the insertion point for the block
                                acBlkTblRec.Origin = new Point3d(startPointX, startPointY, 0);


                                Polyline pl1 = CreatRectangle(startPointX, startPointY, startPointX + 400, startPointY + 287); // Khung bao ngoai
                                Polyline pl2 = CreatRectangle(startPointX + 5, startPointY + 5, startPointX + 395, startPointY + 282); // khung trong
                                acBlkTblRec.AppendEntity(pl1); //them vao block
                                acBlkTblRec.AppendEntity(pl2); //them vao block
                                //cot
                                Line line1, line2;
                                for (int i = 0; i < 7; i++)
                                {
                                    line1 = CreatLine(startPointX + 37.5 + i * 53, startPointY, startPointX + 37.5 + i * 53, startPointY + 5); //vach canh bottom
                                    line2 = CreatLine(startPointX + 37.5 + i * 53, startPointY + 282, startPointX + 37.5 + i * 53, startPointY + 287); //vach canh top 
                                    acBlkTblRec.AppendEntity(line1); //them vao block
                                    acBlkTblRec.AppendEntity(line2); //them vao block
                                }
                                for (int i = 0; i < 5; i++)
                                {
                                    line1 = CreatLine(startPointX, startPointY + 44.5 + i * 49.5, startPointX + 5, startPointY + 44.5 + i * 49.5); //vach canh left
                                    line2 = CreatLine(startPointX + 395, startPointY + 44.5 + i * 49.5, startPointX + 400, startPointY + 44.5 + i * 49.5); //vach canh right
                                    acBlkTblRec.AppendEntity(line1); //them vao block
                                    acBlkTblRec.AppendEntity(line2); //them vao block
                                }
                                Circle c1 = CreatCircle(startPointX - 6, startPointY + 103.5, 3);
                                Circle c2 = CreatCircle(startPointX - 6, startPointY + 183.5, 3);
                                acBlkTblRec.AppendEntity(c1); //them vao block
                                acBlkTblRec.AppendEntity(c2); //them vao block
                                for (int i = 0; i < 8; i++)
                                {
                                    DBText text1 = CreatText(startPointX + 11.25 + i * 52.5, startPointY + 2.55, 2.5, (i + 1).ToString());
                                    DBText text2 = CreatText(startPointX + 11.25 + i * 52.5, startPointY + 284.45, 2.5, (i + 1).ToString());
                                    acBlkTblRec.AppendEntity(text1); //them vao block
                                    acBlkTblRec.AppendEntity(text2); //them vao block
                                }

                                char mychar = 'F';
                                for (int i = 0; i < 6; i++)
                                {
                                    string myString = ((char)(mychar - i)).ToString();
                                    DBText text3 = CreatText(startPointX + 3.18, startPointY + 19.75 + i * 49.5, 2.5, myString);
                                    DBText text4 = CreatText(startPointX + 396.83, startPointY + 19.75 + i * 49.5, 2.5, myString);
                                    acBlkTblRec.AppendEntity(text3); //them vao block
                                    acBlkTblRec.AppendEntity(text4); //them vao block
                                }
                                acBlkTbl.UpgradeOpen();
                                acBlkTbl.Add(acBlkTblRec);
                                acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
                            }
                            acTrans.Commit();
                        }
                        //Dinh nghia lai
                        //else
                        //{
                        //    blkRecId = acBlkTbl[nameBlock];
                        //    BlockTableRecord acBlkTblRec = acTrans.GetObject(blkRecId, OpenMode.ForWrite) as BlockTableRecord;
                        //    acBlkTblRec.Erase(true);
                        //}
                    }
                }
            }
        }
        public static void DefineBlock(BlockTableRecord acBlkTblRec)
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
                        // Open the Block table for read
                        BlockTable acBlkTbl;
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        ObjectId blkRecId = ObjectId.Null;

                        if (!acBlkTbl.Has(acBlkTblRec.Name))    // Neu chua co block border A3
                        {                            
                            {
                                acBlkTbl.UpgradeOpen();                               
                                acBlkTbl.Add(acBlkTblRec);
                                acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
                            }
                            acTrans.Commit();
                        }
                    }
                }
            }
        }        
        public static void InsertBlockBorderA3(double startPointX, double startPointY,double scale)
        {
            string nameBlock = "BorderA3-Duong";
            DefineBlockBorderA3(nameBlock);
            InsertBlock(nameBlock, startPointX, startPointY, scale);
        }
        public static void InsertBlockTitleA3(double startPointX, double startPointY, double scale)
        {
            string nameBlock = "BlockTileA3-Duong";

            DefineBlock(DefineBlockTitleA3(nameBlock));
           InsertBlock(nameBlock, startPointX, startPointY, scale);
        }
        public static void InsertBlock(string blockName,double startPointX, double startPointY,double scale)
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
                        //Get the block definition "Check".                        
                        BlockTable bt = acCurDb.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord blockDef = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;

                        //Also open modelspace - we'll be adding our BlockReference to it
                        BlockTableRecord ms = bt[BlockTableRecord.ModelSpace].GetObject(OpenMode.ForWrite) as BlockTableRecord;

                        //Create new BlockReference, and link it to our block definition
                        Point3d point = new Point3d(startPointX, startPointY, 0);
                        using (BlockReference blockRef = new BlockReference(point, blockDef.ObjectId))
                        {
                            blockRef.ScaleFactors =new Scale3d(scale,scale,scale);
                            //Add the block reference to modelspace
                            ms.AppendEntity(blockRef);
                            acTrans.AddNewlyCreatedDBObject(blockRef, true);

                            //Iterate block definition to find all non-constant
                            // AttributeDefinitions
                            foreach (ObjectId id in blockDef)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForRead);
                                AttributeDefinition attDef = obj as AttributeDefinition;
                                if ((attDef != null) && (!attDef.Constant))
                                {
                                    //This is a non-constant AttributeDefinition
                                    //Create a new AttributeReference
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetDatabaseDefaults();// optional or wrong maybe
                                        attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                        attRef.Position = blockRef.Position.TransformBy(blockRef.BlockTransform);
                                       
                                        //Add the AttributeReference to the BlockReferen
                                        blockRef.AttributeCollection.AppendAttribute(attRef);
                                        acTrans.AddNewlyCreatedDBObject(attRef, true);
                                    }
                                }
                            }
                        }
                        //Our work here is done
                        acTrans.Commit();
                    }
                }
            }    
        }
        private static void AddToModelSpace(Entity ent, Database db)    //Them doi tuong vao database cua ban ve
        {

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace],OpenMode.ForWrite) as BlockTableRecord;
                modelSpace.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent, true);
                trans.Commit();
            }
        }
        public static BlockTableRecord DefineBlockTitleA3(string nameBlockTitle)
        {
            BlockTableRecord acBlkTblRec = new BlockTableRecord();
            acBlkTblRec.Name = nameBlockTitle;
            acBlkTblRec.Origin = new Point3d(0, 0, 0);

            Polyline pl1 = CreatRectangle(0, 0, 180, 47); // Khung bao ngoai
            Line[] arrLine = new Line[14];

            arrLine[0] = CreatLine(0, 8, 180, 8); //Row1
            arrLine[1] = CreatLine(0, 16, 180, 16); //Row2
            arrLine[2] = CreatLine(150, 24, 180, 24); //Row3
            arrLine[3] = CreatLine(0, 32, 180, 32); //Row4
            arrLine[4] = CreatLine(0, 37, 180, 37); //Row5
            arrLine[5] = CreatLine(0, 42, 180, 42); //Row6
            arrLine[6] = CreatLine(10, 32, 10, 47); //Col 1
            arrLine[7] = CreatLine(40, 0, 40, 16); //Col 2
            arrLine[8] = CreatLine(50, 32, 50, 47); //Col 3
            arrLine[9] = CreatLine(80, 0, 80, 32); //Col 4
            arrLine[10] = CreatLine(115, 0, 115,8); //Col 5
            
            arrLine[11] = CreatLine(120, 32, 120, 47); //Col 7
            arrLine[12] = CreatLine(150, 0, 150, 47); //Col 8
            arrLine[13] = CreatLine(165, 0, 165, 24); //Col 9
            int vol = arrLine.Length;
            for (int i = 0; i < vol; i++)
            {
                acBlkTblRec.AppendEntity(arrLine[i]); //them vao block
            }
            acBlkTblRec.AppendEntity(pl1); //them vao block

            TextStyleTable acTextStyleTable = GetTextStyleTable();

            //Ten thiet ke
            MText labelThietKe = new MText();
            labelThietKe.Contents = "Thiết kế:";
            labelThietKe.Location = new Point3d(0.75, 6.33, 0);
            labelThietKe.Attachment = AttachmentPoint.MiddleLeft;
            labelThietKe.TextHeight = 2;
            labelThietKe.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelThietKe);

           
            MText labelDate = new MText();
            labelDate.Contents = DateTime.Now.Month.ToString()+"/"+DateTime.Now.Year.ToString();
            labelDate.Location = new Point3d(21, 6.33, 0);
            labelDate.Attachment = AttachmentPoint.MiddleLeft;
            labelDate.TextHeight = 2;
            labelDate.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelDate);

            MText mtxtThietKe = new MText();
            mtxtThietKe.Contents = @"%<\AcVar CustomDP.Thiết kế>%";
            mtxtThietKe.Location = new Point3d(0.75, 2.5, 0);
            mtxtThietKe.Attachment = AttachmentPoint.MiddleLeft;
            mtxtThietKe.TextHeight = 2.5;
            mtxtThietKe.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtThietKe);          
            
            //Duyet
            MText labelDuyet = new MText();
            labelDuyet.Contents = "Duyệt:";
            labelDuyet.Location = new Point3d(0.75, 14.33, 0);
            labelDuyet.Attachment = AttachmentPoint.MiddleLeft;
            labelDuyet.TextHeight = 2;
            labelDuyet.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelDuyet);

            MText mtxtDuyet = new MText();
            mtxtDuyet.Contents = @"%<\AcVar CustomDP.Duyệt>%";
            mtxtDuyet.Location = new Point3d(0.75, 10.5, 0);
            mtxtDuyet.Attachment = AttachmentPoint.MiddleLeft;
            mtxtDuyet.TextHeight = 2.5;
            mtxtDuyet.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtDuyet);

            //Ban Thiết Kế
            MText labelBTK = new MText();
            labelBTK.Contents = "Ban thiết kế:";
            labelBTK.Location = new Point3d(41, 14.33, 0);
            labelBTK.Attachment = AttachmentPoint.MiddleLeft;
            labelBTK.TextHeight = 2;
            labelBTK.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelBTK);

            MText mtxtBTK = new MText();
            mtxtBTK.Contents = @"%<\AcVar CustomDP.Ban thiết kế>%";
            mtxtBTK.Location = new Point3d(41, 10.5, 0);
            mtxtBTK.Attachment = AttachmentPoint.MiddleLeft;
            mtxtBTK.TextHeight = 2.5;
            mtxtBTK.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtBTK);

            //Ten kiem soat
            MText labelKiemSoat = new MText();
            labelKiemSoat.Contents = "Kiểm soát:";
            labelKiemSoat.Location = new Point3d(41, 6.33, 0);
            labelKiemSoat.Attachment = AttachmentPoint.MiddleLeft;
            labelKiemSoat.TextHeight = 2;
            labelKiemSoat.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelKiemSoat);            

            MText mtxtKiemSoat = new MText();
            mtxtKiemSoat.Contents = @"%<\AcVar CustomDP.Kiểm soát>%";
            mtxtKiemSoat.Location = new Point3d(41, 2.5, 0);
            mtxtKiemSoat.Attachment = AttachmentPoint.MiddleLeft;
            mtxtKiemSoat.TextHeight = 2.5;
            mtxtKiemSoat.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtKiemSoat);

            //Ten Ma so
            MText labelMaSo = new MText();
            labelMaSo.Contents = "Mã số:";
            labelMaSo.Location = new Point3d(81.75, 6.33, 0);
            labelMaSo.Attachment = AttachmentPoint.MiddleLeft;
            labelMaSo.TextHeight = 2;
            labelMaSo.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelMaSo);

            MText mtxtMaSo = new MText();
            mtxtMaSo.Contents = @"%<\AcVar CustomDP.MS>%";
            mtxtMaSo.Location = new Point3d(81.75, 2.5, 0);
            mtxtMaSo.Attachment = AttachmentPoint.MiddleLeft;
            mtxtMaSo.TextHeight = 2.5;
            mtxtMaSo.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtMaSo);

            //Ma so ban ve
            MText labelMaSoBanVe = new MText();
            labelMaSoBanVe.Contents = "Số bản vẽ:";
            labelMaSoBanVe.Location = new Point3d(115.75, 6.33, 0);
            labelMaSoBanVe.Attachment = AttachmentPoint.MiddleLeft;
            labelMaSoBanVe.TextHeight = 2;
            labelMaSoBanVe.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelMaSoBanVe);

            //Phần đánh mã số bản vẽ
            MText mtxtMSBV = new MText();
            mtxtMSBV.Contents = @"%<\AcVar CustomDP.Công suất>%D-";
            mtxtMSBV.Location = new Point3d(127, 2.5, 0);
            mtxtMSBV.Attachment = AttachmentPoint.MiddleRight;
            mtxtMSBV.TextHeight = 2.5;
            mtxtMSBV.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtMSBV);

            AttributeDefinition attMSBV = new AttributeDefinition();
            attMSBV.Justify = AttachmentPoint.MiddleLeft;
            attMSBV.Position = new Point3d(127.5, 2.5, 0);
            attMSBV.AlignmentPoint = attMSBV.Position;           
            attMSBV.Preset = true;
            attMSBV.LockPositionInBlock = true;
            attMSBV.Verifiable = true;
            attMSBV.Prompt = "Mã số bản vẽ";
            attMSBV.Height = 2.5;
            attMSBV.Tag = "MSBV";
            attMSBV.TextString = "01-04-000";
            attMSBV.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(attMSBV);


            //Tỷ lệ BV
            MText labelTyLe = new MText();
            labelTyLe.Contents = "Tỷ lệ:";
            labelTyLe.Location = new Point3d(151, 6.33, 0);
            labelTyLe.Attachment = AttachmentPoint.MiddleLeft;
            labelTyLe.TextHeight = 2;
            labelTyLe.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelTyLe);

            AttributeDefinition attTyLeBV = new AttributeDefinition();
            attTyLeBV.Justify = AttachmentPoint.MiddleLeft;
            attTyLeBV.AlignmentPoint = new Point3d(151, 2.5, 0);
            attTyLeBV.Preset = true;
            attTyLeBV.LockPositionInBlock = true;
            attTyLeBV.Verifiable = true;
            attTyLeBV.Prompt = "Tỷ lệ bản vẽ";
            attTyLeBV.Height = 2.5;
            attTyLeBV.Tag = "TyLeBV";
            attTyLeBV.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(attTyLeBV);

            //Tờ số
            MText labelToSo = new MText();
            labelToSo.Contents = "Tờ số:";
            labelToSo.Location = new Point3d(151, 14.33, 0);
            labelToSo.Attachment = AttachmentPoint.MiddleLeft;
            labelToSo.TextHeight = 2;
            labelToSo.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelToSo);

            AttributeDefinition attToSo = new AttributeDefinition();
            attToSo.Justify = AttachmentPoint.MiddleLeft;
            attToSo.AlignmentPoint = new Point3d(151, 10.5, 0);            
            attToSo.Preset = true;
            attToSo.LockPositionInBlock = true;
            attToSo.Verifiable = true;
            attToSo.Prompt = "Số thứ tự bản vẽ";
            attToSo.Height = 2.5;
            attToSo.Tag = "SoTTBV";
            attToSo.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(attToSo);


            //Xuất sứ
            MText labelXuatXu = new MText();
            labelXuatXu.Contents = "Xuất xứ:";
            labelXuatXu.Location = new Point3d(166, 14.33, 0);
            labelXuatXu.Attachment = AttachmentPoint.MiddleLeft;
            labelXuatXu.TextHeight = 2;
            labelXuatXu.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelXuatXu);

            AttributeDefinition attXuatXu = new AttributeDefinition();
            attXuatXu.Justify = AttachmentPoint.MiddleLeft;
            attXuatXu.AlignmentPoint = new Point3d(166, 10.5, 0);
            attXuatXu.Preset = true;
            attXuatXu.LockPositionInBlock = true;
            attXuatXu.Verifiable = true;
            attXuatXu.Prompt = "Xuất xứ";
            attXuatXu.TextString = "B3";
            attXuatXu.Height = 2.5;
            attXuatXu.Tag = "XuatXu";
            attXuatXu.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(attXuatXu);            

            //Tên ban ve
            MText labelTenBV = new MText();
            labelTenBV.Contents = "Tên bản vẽ:";
            labelTenBV.Location = new Point3d(81, 14.33, 0);
            labelTenBV.Attachment = AttachmentPoint.MiddleLeft;
            labelTenBV.TextHeight = 2;
            labelTenBV.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelTenBV);

            AttributeDefinition attTenBV = new AttributeDefinition();
            attTenBV.Justify = AttachmentPoint.MiddleCenter;
            attTenBV.AlignmentPoint = new Point3d(115, 11, 0);            
            attTenBV.Preset = true;
            attTenBV.LockPositionInBlock = true;
            attTenBV.Verifiable = true;
            attTenBV.Prompt = "Tên bản vẽ";
            attTenBV.Tag = "TenBV";
            attTenBV.TextString = "MẠCH TỪ";
            attTenBV.Height = 3;
            attTenBV.TextStyleId = acTextStyleTable["X1_BTK"];
            attTenBV.IsMTextAttributeDefinition = true;            
            acBlkTblRec.AppendEntity(attTenBV);

            //Số lượng
            MText labelSoLuong = new MText();
            labelSoLuong.Contents = "Số lượng:";
            labelSoLuong.Location = new Point3d(151, 22, 0);
            labelSoLuong.Attachment = AttachmentPoint.MiddleLeft;
            labelSoLuong.TextHeight = 2;
            labelSoLuong.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelSoLuong);

            AttributeDefinition attSoLuong = new AttributeDefinition();
            attSoLuong.Justify = AttachmentPoint.MiddleLeft;
            attSoLuong.AlignmentPoint = new Point3d(151, 18.5, 0);
            attSoLuong.Preset = true;
            attSoLuong.LockPositionInBlock = true;
            attSoLuong.Verifiable = true;
            attSoLuong.Prompt = "Số lượng";
            attSoLuong.Tag = "SoLuong";
            attSoLuong.Height = 2.5;
            attSoLuong.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(attSoLuong);

            //Khối lượng
            MText labelKhoiLuong = new MText();
            labelKhoiLuong.Contents = "K.lượng:";
            labelKhoiLuong.Location = new Point3d(166, 22, 0);
            labelKhoiLuong.Attachment = AttachmentPoint.MiddleLeft;
            labelKhoiLuong.TextHeight = 2;
            labelKhoiLuong.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelKhoiLuong);

            AttributeDefinition attKhoiLuong = new AttributeDefinition();
            attKhoiLuong.Justify = AttachmentPoint.MiddleLeft;
            attKhoiLuong.AlignmentPoint = new Point3d(166, 18.5, 0);
            attKhoiLuong.Preset = true;
            attKhoiLuong.LockPositionInBlock = true;
            attKhoiLuong.Verifiable = true;
            attKhoiLuong.Prompt = "Khối lượng";
            attKhoiLuong.Tag = "KhoiLuong";
            attKhoiLuong.Height = 2.5;
            attKhoiLuong.TextStyleId = acTextStyleTable["X1_BTK"];
            attTenBV.IsMTextAttributeDefinition = true;
            acBlkTblRec.AppendEntity(attKhoiLuong);

            //Tên máy
            MText labelTenMay = new MText();
            labelTenMay.Contents = "Tên máy:";
            labelTenMay.Location = new Point3d(81, 30, 0);
            labelTenMay.Attachment = AttachmentPoint.MiddleLeft;
            labelTenMay.TextHeight = 2;
            labelTenMay.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelTenMay);

            MText mtxtTenMay = new MText();
            mtxtTenMay.Contents = @"MBA %<\AcVar CustomDP.Công suất>%VA-%<\AcVar CustomDP.Điện áp>%kV \P TRẠM: %<\AcVar CustomDP.Trạm>%";
            mtxtTenMay.Location = new Point3d(115, 24, 0);
            mtxtTenMay.Attachment = AttachmentPoint.MiddleCenter;
            mtxtTenMay.TextHeight = 2.5;
            mtxtTenMay.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtTenMay);

            //Tên + Địa chỉ + logo công ty
            MText mtxtTenCTy = new MText();
            mtxtTenCTy.Contents = @"TỔNG CÔNG TY \P THIẾT BỊ ĐIỆN ĐÔNG ANH \P Địa chỉ: Số 189 - Đường Lâm Tiên - TT Đông Anh - Hà Nội \P Tel: (84.4) 38833779  Fax: (84.4)38833113";
            mtxtTenCTy.Location = new Point3d(45.5, 24, 0);
            mtxtTenCTy.Attachment = AttachmentPoint.MiddleCenter;
            mtxtTenCTy.TextHeight = 1.8;
            mtxtTenCTy.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(mtxtTenCTy);

            // Get the image dictionary
            

            //Vật liệu
            MText labelVatLieu = new MText();
            labelVatLieu.Contents = "Vật liệu:";
            labelVatLieu.Location = new Point3d(151, 30, 0);
            labelVatLieu.Attachment = AttachmentPoint.MiddleLeft;
            labelVatLieu.TextHeight = 2;
            labelVatLieu.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelVatLieu);

            AttributeDefinition attVatLieu = new AttributeDefinition();
            attVatLieu.Justify = AttachmentPoint.MiddleLeft;
            attVatLieu.AlignmentPoint = new Point3d(151, 26.5, 0);
            attVatLieu.Preset = true;
            attVatLieu.LockPositionInBlock = true;
            attVatLieu.Verifiable = true;
            attVatLieu.Prompt = "Vật liệu";
            attVatLieu.Tag = "VatLieu";
            attVatLieu.Height = 2.5;
            attVatLieu.TextStyleId = acTextStyleTable["X1_BTK"];
            attXuatXu.IsMTextAttributeDefinition = true;
            acBlkTblRec.AppendEntity(attVatLieu);

            //Số sửa đổi
            MText labelSoSuaDoi = new MText();
            labelSoSuaDoi.Contents = "SSĐ";
            labelSoSuaDoi.Location = new Point3d(5, 34.5, 0);
            labelSoSuaDoi.Attachment = AttachmentPoint.MiddleCenter;
            labelSoSuaDoi.TextHeight = 2.5;
            labelSoSuaDoi.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelSoSuaDoi);
            //Tên sửa đổi
            MText labelTenSuaDoi = new MText();
            labelTenSuaDoi.Contents = "Tên";
            labelTenSuaDoi.Location = new Point3d(30, 34.5, 0);
            labelTenSuaDoi.Attachment = AttachmentPoint.MiddleCenter;
            labelTenSuaDoi.TextHeight = 2.5;
            labelTenSuaDoi.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelTenSuaDoi);
            //Nội dung sửa đổi
            MText labelNoiDungSuaDoi = new MText();
            labelNoiDungSuaDoi.Contents = "Nội dung sửa đổi";
            labelNoiDungSuaDoi.Location = new Point3d(85, 34.5, 0);
            labelNoiDungSuaDoi.Attachment = AttachmentPoint.MiddleCenter;
            labelNoiDungSuaDoi.TextHeight = 2.5;
            labelNoiDungSuaDoi.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelNoiDungSuaDoi);
            //Ngày sửa đổi
            MText labelNgaySuaDoi = new MText();
            labelNgaySuaDoi.Contents = "Ngày";
            labelNgaySuaDoi.Location = new Point3d(135, 34.5, 0);
            labelNgaySuaDoi.Attachment = AttachmentPoint.MiddleCenter;
            labelNgaySuaDoi.TextHeight = 2.5;
            labelNgaySuaDoi.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelNgaySuaDoi);
            //Chứ ký sửa đổi
            MText labelKySuaDoi = new MText();
            labelKySuaDoi.Contents = "Chữ ký";
            labelKySuaDoi.Location = new Point3d(165, 34.5, 0);
            labelKySuaDoi.Attachment = AttachmentPoint.MiddleCenter;
            labelKySuaDoi.TextHeight = 2.5;
            labelKySuaDoi.TextStyleId = acTextStyleTable["X1_BTK"];
            acBlkTblRec.AppendEntity(labelKySuaDoi);            

            //MText mtxtTenBV = new MText();
            //mtxtTenBV.Contents = "{\\fTCVN 7284|b1|i0|c0|p34;MẠCH TỪ}";
            //mtxtTenBV.Location = new Point3d(115, 10.5, 0);
            //mtxtTenBV.Attachment = AttachmentPoint.MiddleCenter;
            //mtxtTenBV.TextHeight = 3;
            //mtxtTenBV.TextStyleId = acTextStyleTable["X1_BTK"];
            //acBlkTblRec.AppendEntity(mtxtTenBV);
            string path = @"C:\Logo.png";
            if (!File.Exists(@"C:\Logo.png"))
            {
                OpenFileDialog opFileDialog1 = new OpenFileDialog();
                opFileDialog1.Filter = "All file|*.*";
                opFileDialog1.Title = "Load Logo Image";
                if (opFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    path = opFileDialog1.FileName;
                }                
            }
            DefineRasterImage("logo", path);
            RasterImage imgLogo = CreatRasterImage(1.8, 19.2, "logo");
            acBlkTblRec.AppendEntity(imgLogo);            
                       
            return acBlkTblRec;
        }
        public static void DefineRasterImage(string imgName,string imgPath)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            using (DocumentLock acLockDoc = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    #region "define IMG"
                    ObjectId acImgDefId;

                    // Get the image dictionary
                    ObjectId acImgDctID = RasterImageDef.GetImageDictionary(acCurDb);

                    // Check to see if the dictionary does not exist, it not then create it
                    if (acImgDctID.IsNull)
                    {
                        acImgDctID = RasterImageDef.CreateImageDictionary(acCurDb);
                    }

                    // Open the image dictionary
                    DBDictionary acImgDict = acTrans.GetObject(acImgDctID, OpenMode.ForRead) as DBDictionary;

                    // Check to see if the image definition already exists
                    if (acImgDict.Contains(imgName))
                    {
                        //Nếu đã có thì không tạo mới
                    }
                    else
                    {
                        // Create a raster image definition
                        RasterImageDef acRasterDefNew = new RasterImageDef();

                        // Set the source for the image file
                        acRasterDefNew.SourceFileName = imgPath;

                        // Load the image into memory
                        acRasterDefNew.Load();

                        // Add the image definition to the dictionary
                        acImgDict.UpgradeOpen();
                        acImgDefId = acImgDict.SetAt(imgName, acRasterDefNew);
                        acTrans.AddNewlyCreatedDBObject(acRasterDefNew, true);                        
                    }                    
                    acTrans.Commit();
                    #endregion
                }
            } 
        }
        public static RasterImage CreatRasterImage(double startPointX, double startPointY,string nameImg)
        {
            
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (acDoc != null)
            {
                Editor acEdit = acDoc.Editor;
                Database acCurDb = acDoc.Database;
                try
                {
                    using (DocumentLock acLockDoc = acDoc.LockDocument())
                    {
                        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                        {
                            DBDictionary namedDic = (DBDictionary)acTrans.GetObject(acCurDb.NamedObjectsDictionaryId, OpenMode.ForRead);
                            if (!namedDic.Contains("ACAD_IMAGE_DICT"))
                            {
                                acEdit.WriteMessage("\nACAD_IMAGE_DICT does not Exist");
                                return null;
                            }
                            ObjectId dictId = RasterImageDef.GetImageDictionary(acCurDb);
                            DBDictionary dict = (DBDictionary)acTrans.GetObject(dictId, OpenMode.ForRead);

                            if (dict.Contains(nameImg))
                            {
                                acEdit.WriteMessage("\nDoor RasterImage Definition Exists");
                                ObjectId defId = dict.GetAt(nameImg);
                                RasterImage ri = new RasterImage();
                                ri.ImageDefId = defId;                                
                                ri.ShowImage = true;
                                Matrix3d ucs = acEdit.CurrentUserCoordinateSystem;
                                Point3d pt =new Point3d (startPointX,startPointY,0);
                                Vector3d xAxis = new Vector3d(10, 0, 0);
                                Vector3d yAxis = new Vector3d(0, 10, 0);
                                ri.Orientation = new CoordinateSystem3d(pt.TransformBy(ucs), xAxis.TransformBy(ucs), yAxis.TransformBy(ucs));
                                ri.SetClipBoundaryToWholeImage();
                                return ri;
                            }
                            else return null;
                        }
                    }
                }
                catch (System.Exception ex)
                {                    
                	Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.Message );
                    return null;
                }                
            }
            else return null;
        }
        public static DBText CreatText(double startPointX, double startPointY, double height, string content) // chen text vao ban ve
        {
            DBText text = new DBText();
            Point3d pt = new Point3d(startPointX, startPointY, 0);            
            text.TextString = content;
            text.Height = height;            
            text.VerticalMode = TextVerticalMode.TextVerticalMid;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = pt; 
            return text;                     
        }  
        public static void CreatTextStyle(string nameTextStyle, string nameFont,double textSize)
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
                            Autodesk.AutoCAD.GraphicsInterface.FontDescriptor newFont = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(oldFont.TypeFace,false , false, oldFont.CharacterSet, oldFont.PitchAndFamily);
                            newTextStyleTableRecord.Font = newFont;
                            newTextStyleTableRecord.TextSize = textSize;
                            newTextStyleTableRecord.Name = nameTextStyle;
                            newTextStyleTable.Add(newTextStyleTableRecord);
                            acTrans.AddNewlyCreatedDBObject(newTextStyleTableRecord,true);
                        }
                        // Save the new object to the database
                        acTrans.Commit();
                    }
                }
            }
        }
        public static void CreatDimStyle(string nameDimStyle, string textStyle) //Text Heigh = dimScale x 2.5
        {
            TextStyleTable acTextStyleTable = GetTextStyleTable();
            if (!acTextStyleTable.Has(textStyle))
            {
                MessageBox.Show("Chưa khởi tạo:"+textStyle);
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
                        TextStyleTableRecord txtStyleRecord = acTrans.GetObject(acTextStyleTable[textStyle],OpenMode.ForRead) as TextStyleTableRecord;
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
                            newDimStyleTableRecord.Dimgap = txtStyleRecord.TextSize/3;// Offset from Dim Line

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
        public static TextStyleTable GetTextStyleTable()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            TextStyleTable newTextStyleTable=null;
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
        public static void CreatCostumProperties(string key, string value)
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
    }
   
}
