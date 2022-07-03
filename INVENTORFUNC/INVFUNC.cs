using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Windows.Forms;

namespace ACADTRANSFORMER.INVENTORFUNC
{
   public class INVFUNC
    {
        public PlanarSketch InsertPackage(Inventor.Application m_inventorApp,PlanarSketch sketch1, PartDocument opartDoc, PartComponentDefinition oComDef, TransientGeometry oTG,double c_cs, double h_cs,double bt,double bg, double hbg,double bi,SketchPoint centerPoint, bool huongextrude)
        {
            c_cs = c_cs / 10;
            h_cs = h_cs / 10;
            bt = bt / 10;
            bg = bg / 10;
            hbg = hbg / 10;
            bi = bi / 20;
            Point2d tam2d = oTG.CreatePoint2d(0, 0);  // Tam  mach tu tren sket

            //================================Bat dau ve=======================================//
            //----------------------------------Ve duong bao cap ton---------------------------//
            Point2d gocbao = oTG.CreatePoint2d(-(c_cs + bt / 2), -(h_cs + 2 * bg + 2 * hbg) / 2);  // goc duoi ben trai cua hinh vuong
            SketchEntitiesEnumerator hinhvuong1 = sketch1.SketchLines.AddAsTwoPointCenteredRectangle(tam2d, gocbao); //ve hinh vuong
            SketchLine canhtren = (SketchLine)hinhvuong1[1];
            SketchLine canhphai = (SketchLine)hinhvuong1[2];
            SketchLine canhduoi = (SketchLine)hinhvuong1[3];
            SketchLine canhtrai = (SketchLine)hinhvuong1[4];
            sketch1.DimensionConstraints.AddOffset(canhtren, (SketchEntity)canhduoi, tam2d, false); //tao rang buoc giua canh tren va canh duoi
            sketch1.DimensionConstraints.AddOffset(canhtrai, (SketchEntity)canhphai, tam2d, false); //tao rang buoc giua canh tren va canh duoi
            sketch1.DimensionConstraints.AddOffset(canhtren, (SketchEntity)centerPoint, tam2d, false); //tao rang buoc giua canh tren goc toa do
            sketch1.DimensionConstraints.AddOffset(canhtrai, (SketchEntity)centerPoint, tam2d, false); //tao rang buoc giua canh trai goc toa do
                                                                                                       //-----------------Ve cua so mach tu ben trai---------//                   
            Point2d goc_cuasotrai = oTG.CreatePoint2d(-(c_cs - bt / 2), -(h_cs + 2 * hbg) / 2);  // goc duoi ben trai cua hinh vuong
            Point2d tam_cuasotrai = oTG.CreatePoint2d(-c_cs / 2, 0);
            SketchEntitiesEnumerator hinhvuong2 = sketch1.SketchLines.AddAsTwoPointCenteredRectangle(tam_cuasotrai, goc_cuasotrai); //ve hinh vuong
            SketchLine canhtren_cuasotrai = (SketchLine)hinhvuong2[1];
            SketchLine canhphai_cuasotrai = (SketchLine)hinhvuong2[2];
            SketchLine canhduoi_cuasotrai = (SketchLine)hinhvuong2[3];
            SketchLine canhtrai_cuasotrai = (SketchLine)hinhvuong2[4];
            sketch1.DimensionConstraints.AddOffset(canhtren_cuasotrai, (SketchEntity)canhtren, tam2d, false); //tao rang buoc giua gong tren
            sketch1.DimensionConstraints.AddOffset(canhduoi_cuasotrai, (SketchEntity)canhduoi, tam2d, false); //tao rang buoc giua gond uoi
            sketch1.DimensionConstraints.AddOffset(canhtrai_cuasotrai, (SketchEntity)canhtrai, tam2d, false); //tao rang buoc giua canh tren va canh duoi
            sketch1.DimensionConstraints.AddOffset(canhphai_cuasotrai, (SketchEntity)centerPoint, tam2d, false); //tao rang buoc giua canh tren va canh duoi

            //-----------------Ve cua so mach tu ben phai---------//                   
            Point2d goc_cuasophai = oTG.CreatePoint2d(bt / 2, -(h_cs + 2 * hbg) / 2);  // goc duoi ben trai cua hinh vuong
            Point2d tam_cuasophai = oTG.CreatePoint2d(+c_cs / 2, 0);
            SketchEntitiesEnumerator hinhvuong3 = sketch1.SketchLines.AddAsTwoPointCenteredRectangle(tam_cuasophai, goc_cuasophai); //ve hinh vuong
            SketchLine canhtren_cuasophai = (SketchLine)hinhvuong3[1];
            SketchLine canhphai_cuasophai = (SketchLine)hinhvuong3[2];
            SketchLine canhduoi_cuasophai = (SketchLine)hinhvuong3[3];
            SketchLine canhtrai_cuasophai = (SketchLine)hinhvuong3[4];
            sketch1.DimensionConstraints.AddOffset(canhtren_cuasophai, (SketchEntity)canhtren, tam2d, false); //tao rang buoc giua gong tren
            sketch1.DimensionConstraints.AddOffset(canhduoi_cuasophai, (SketchEntity)canhduoi, tam2d, false); //tao rang buoc giua gond uoi
            sketch1.DimensionConstraints.AddOffset(canhphai_cuasophai, (SketchEntity)canhphai, tam2d, false); //tao rang buoc giua canh tren va canh duoi
            sketch1.DimensionConstraints.AddOffset(canhtrai_cuasophai, (SketchEntity)centerPoint, tam2d, false); //tao rang buoc giua canh tren va canh duoi
            Profile profile1 = sketch1.Profiles.AddForSolid();
            ExtrudeFeature extrude1;
            if (huongextrude == true)
            {
                extrude1 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(profile1, bi, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
            }     
            else extrude1 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(profile1, bi, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, PartFeatureOperationEnum.kJoinOperation);
            PlanarSketch ret = oComDef.Sketches.Add(extrude1.EndFaces[1], false);
            return ret;
        }
        public void insertCore(double c_cs, double h_cs, double[] bt,double[] bg, double[] hbg, double[] bi,int[] thongdaucapton, double hr)
        {
            bi[0] = bi[0] + thongdaucapton[0] * hr;
            for (int i = 1; i < bi.Length; i++)
            {
                bi[i] = bi[i] + 2*thongdaucapton[i]*hr;
            }
            Inventor.Application m_inventorApp = null;
            bool m_quitInventor = false;
            try
            {
                try
                {
                    m_inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
                }
                catch
                {
                }
                // If not active, create a new inventor session
                if (m_inventorApp != null) //Inventor da duoc mo
                {
                    //Create new part document
                    PartDocument opartDoc = (PartDocument)m_inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, m_inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject, SystemOfMeasureEnum.kMetricSystemOfMeasure));
                    PartComponentDefinition oComDef = (PartComponentDefinition)opartDoc.ComponentDefinition;
                    TransientGeometry oTG = m_inventorApp.TransientGeometry;

                    //Khai bao cac bien nguoi dung
                    oComDef.Parameters.UserParameters.AddByExpression("Ci", "1480", UnitsTypeEnum.kMillimeterLengthUnits);
                    oComDef.Parameters.UserParameters.AddByExpression("Bti", "240", UnitsTypeEnum.kMillimeterLengthUnits);
                    oComDef.Parameters.UserParameters.AddByExpression("Bgi", "340", UnitsTypeEnum.kMillimeterLengthUnits);
                    oComDef.Parameters.UserParameters.AddByExpression("Hi", "340", UnitsTypeEnum.kMillimeterLengthUnits);

                    
                    
                    Point2d tam2d = oTG.CreatePoint2d(0, 0);  // Tam  mach tu tren sket
                    // Ve cap dau tien
                    List<PlanarSketch> lstSket1 = new List<PlanarSketch>(); //mang sket nua duong
                    List<PlanarSketch> lstSket2 = new List<PlanarSketch>(); //mang sket nua am
                    //Tao Sketc1
                    PlanarSketch sketch1 = oComDef.Sketches.AddWithOrientation(oComDef.WorkPlanes[3],oComDef.WorkAxes[1], true, true, oComDef.WorkPoints[1]);
                    PlanarSketch sketch2 = oComDef.Sketches.AddWithOrientation(oComDef.WorkPlanes[3],oComDef.WorkAxes[1], false,true, oComDef.WorkPoints[1]);
                    SketchPoint centerPoint = (SketchPoint)sketch1.AddByProjectingEntity(oComDef.WorkAxes[3]);
                    lstSket1.Add(sketch1);
                    lstSket2.Add(sketch2);
                    int i = 0;
                    while(i<bt.Length)
                    {
                        if(bt[i] != 0 && bg[i] != 0 && bi[i] != 0)
                        {
                            
                            lstSket1.Add(InsertPackage(m_inventorApp, lstSket1[i], opartDoc, oComDef, oTG, c_cs, h_cs, bt[i], bg[i], hbg[i], bi[i], centerPoint, true));                            
                        }
                        i = i + 1;
                    }
                    int j = 0;
                    while (j < bt.Length)
                    {                        
                        if (bt[j] != 0 && bg[j] != 0 && bi[j] != 0)
                        {
                            if(j==0)
                            {
                                lstSket2.Add(InsertPackage(m_inventorApp, lstSket2[j], opartDoc, oComDef, oTG, c_cs, h_cs, bt[j], bg[j], hbg[j], bi[j], centerPoint, false));                                
                            }                            
                            else lstSket2.Add(InsertPackage(m_inventorApp, lstSket2[j], opartDoc, oComDef, oTG, c_cs, h_cs, bt[j], bg[j], hbg[j], bi[j], centerPoint, true));                            
                        }
                        j = j + 1;
                    }
                }
                else //Inventor chua duoc mo
                {
                    MessageBox.Show("Bạn phải mở inventor lên đã");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem getting some Inventor information", "Error", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                if (m_inventorApp != null && m_quitInventor == true) m_inventorApp.Quit();
                m_inventorApp = null;
            }
        }
    }
    
}

