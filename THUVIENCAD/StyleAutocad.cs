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
    class StyleAutocad
    {
        //public static void CreatDefaultLayer()
        //{
        //    Tao cac layer co ban
        //    Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        //    Editor ed = doc.Editor;
        //    Database db = doc.Database;
        //    string visibleLayer = "Visible";
        //    using (DocumentLock acLockDoc = doc.LockDocument())
        //    {
        //        using (Transaction acTrans = db.TransactionManager.StartTransaction())
        //        {
        //            LayerTable lt = (LayerTable)acTrans.GetObject(db.LayerTableId, OpenMode.ForRead);
        //            SymbolUtilityServices.ValidateSymbolName(visibleLayer, false);
        //            if (!lt.Has(visibleLayer))
        //            {
        //                 Create our new layer table record...
        //                 ... and set its properties
        //                LayerTableRecord ltr = new LayerTableRecord();
        //                ltr.Name = visibleLayer;
        //                ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
        //                ltr.LineWeight = lineWeight;
        //                ltr.IsPlottable = isPlottAble;
        //                ltr.lin
        //                 Add the new layer to the layer table
        //                lt.UpgradeOpen();
        //                ObjectId ltId = lt.Add(ltr);
        //                acTrans.AddNewlyCreatedDBObject(ltr, true);
        //            }
        //            else MessageBox.Show("\nA layer with this name already exists.");
        //            acTrans.Commit();
        //        }
        //    }
        //}
        public static void LoadLineType(string sLineTypName)
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
                    acLineTypTbl = acTrans.GetObject(db.LinetypeTableId,OpenMode.ForRead) as LinetypeTable;
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
        public static void CreatLayer(string layerName, string sLineTypName, short colorIndex, LineWeight lineWeight, bool isPlottAble)
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

                    SymbolUtilityServices.ValidateSymbolName(layerName,false);
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
    }
}
