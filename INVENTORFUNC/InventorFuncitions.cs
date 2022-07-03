using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Windows.Forms;

namespace ACADTRANSFORMER.INVENTORFUNC
{
    public class InventorFuncitions
    {
        #region "propeties"      
        private double _c;             //Khoảng cách hai tâm trụ
        private double _hcs;           //Chiều cao cửa sổ mạch từ
        private double _hr = 4;        //Bề dày que thông dầu
        private double _delta = 0.27;  //Bề dày lá tôn
        private double _steplab = 10;  //bước lệch đỉnh
        
        private double _ke = 0.97;        //Hệ số suy giảm bề dày tôn
        double[] _thongdau = new double[19];
        private double[] _bt = new double[19];  //Cấp trụ
        private double[] _bg = new double[19];  //Cấp gông
        private double[] _hbg = new double[19]; //ha bac gong
        private double[] _offsetHB = new double[19]; //
        private int[] _thongdaucapton = new int[19]; //

        private int[] _n = new int[19];         // Số lá từng cấp
        private double[] _beday = new double[19];
        #endregion

        public void DrawingCore(double c_cs, double h_cs )
        {                       
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
                //If not active, create a new inventor session
                if (m_inventorApp != null) //Inventor da duoc mo
                {
                    //Create new part document
                    PartDocument opartDoc = (PartDocument)m_inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, m_inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject, SystemOfMeasureEnum.kMetricSystemOfMeasure));
                    PartComponentDefinition oComDef = (PartComponentDefinition)opartDoc.ComponentDefinition;
                    TransientGeometry oTG = m_inventorApp.TransientGeometry;
                                 
                    //Khoi tao goc toa do
                    Inventor.Point oPoint = oTG.CreatePoint(0, 0, 0);
                    //Khoi tao vec to don vi z;
                    Inventor.UnitVector zVector = oTG.CreateUnitVector(0, 0, 1);

                    //Tai tryc axisZ theo phuong Z  
                    WorkAxis axisZ = oComDef.WorkAxes.AddFixed(oPoint, zVector);
                    axisZ.Name = "Truc Z";
                    axisZ.Visible = false;

                    //Tao plane1
                    WorkPlane plane1 = oComDef.WorkPlanes.AddFixed(oTG.CreatePoint(0, 0, 0), oTG.CreateUnitVector(1, 0, 0), oTG.CreateUnitVector(0, 1, 0));

                    //Dung plane1 de tao sketch1
                    PlanarSketch sketch1 = oComDef.Sketches.Add(plane1);
                    SketchPoint centerPoint = (SketchPoint)sketch1.AddByProjectingEntity(axisZ);

                    

                    ////---------------------Ve trụ giữa----------//
                    //SketchPoint trugiua_p1 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - lechdinh, y0 + h_trugiua / 2), false);        //Đỉnh trên
                    //SketchPoint trugiua_p2 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + lechdinh, y0 - h_trugiua / 2), false); //Đỉnh dưới
                    //SketchPoint trugiua_p3 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - bt / 2, y0 + h_trugiua / 2 - bt / 2 + lechdinh), false); //Đỉnh trên bên trái
                    //SketchPoint trugiua_p4 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + bt / 2, y0 + h_trugiua / 2 - bt / 2 - lechdinh), false); //Đỉnh trên bên phải
                    //SketchPoint trugiua_p5 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - bt / 2, y0 - h_trugiua / 2 + bt / 2 + lechdinh), false); //Đỉnh dưới bên trái
                    //SketchPoint trugiua_p6 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + bt / 2, y0 - h_trugiua / 2 + bt / 2 - lechdinh), false); //Đỉnh dưới bên phải
                    //SketchLine trugiua_ltrentrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p1, trugiua_p3);
                    //SketchLine trugiua_ltrenphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p1, trugiua_p4);
                    //SketchLine trugiua_ltrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p3, trugiua_p5);
                    //SketchLine trugiua_lduoiphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p2, trugiua_p6);
                    //SketchLine trugiua_lphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p4, trugiua_p6);
                    //SketchLine trugiua_lduoitrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p5, trugiua_p2);

                   
                }
                else //Inventor chua duoc mo
                {
                    //Inventor chua duoc mo
                    MessageBox.Show("Bạn phải mở inventor lên đã");
                    //Type inventorAppType = Type.GetTypeFromProgID("Inventor.Application");
                    //m_inventorApp = Activator.CreateInstance(inventorAppType) as Inventor.Application;
                    //m_quitInventor = true;
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


        /// <summary>
        /// Hàm vẽ một cấp tôn nào đó
        /// </summary>
        /// <param name="ccs">Khoảng cách tâm trụ</param>
        /// <param name="hcs">chiều cao cửa số cấp tôn</param>
        /// <param name="bt">khổ tôn trụ</param>
        /// <param name="bg">khổ tôn gông</param>
        /// <param name="ltruben">chiều cao lá tôn trụ bên</param>
        /// <param name="lg">chiều dài lá gông</param>
        /// <param name="d">chiều dày lá tôn</param>
        /// <param name="n">số lá tôn</param>
        /// <param name="steplap">chieu dai lech dinh</param>
        /// <param name="sodinh">số đỉnh</param>
        //public void DrawingPackage(double ccs, double hcs, double bt, double bg, double ltruben, double lg, double d, double sola, double steplap, int sodinh, int soLaPerStep) //Ve mot bac ton
        public void DrawingPackage() //Ve mot bac ton
        {
            double ccs=1480;
            double hcs=1570;
            double bt=680;
            double bg = 680;
            double lg = 1480;
            double d = 0.23; // be day la ton
            int sola = 100; //so la ton
            double steplap = 10; // chieu dai lech dinh
            int soLaPerStep = 2;//So la mot dinh;
            int sodinhPerPack = 5;//So dinh (4,5,6)
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
                //If not active, create a new inventor session
                if (m_inventorApp != null) //Inventor da duoc mo
                {
                    //Create new part document
                    PartDocument opartDoc = (PartDocument)m_inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, m_inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject, SystemOfMeasureEnum.kMetricSystemOfMeasure));
                    PartComponentDefinition oComDef = (PartComponentDefinition)opartDoc.ComponentDefinition;
                    TransientGeometry oTG = m_inventorApp.TransientGeometry;
                    
                    Inventor.Point oPoint = oTG.CreatePoint(0, 0, 0); //Khoi tao goc toa do                   
                    Inventor.UnitVector zVector = oTG.CreateUnitVector(0, 0, 1); //Khoi tao vec to don vi z;                      
                    WorkAxis axisZ = oComDef.WorkAxes.AddFixed(oPoint, zVector); //Tai tryc axisZ theo phuong Z  
                    axisZ.Name = "Truc Z";
                    axisZ.Visible = false;                   
                    WorkPlane plane1 = oComDef.WorkPlanes.AddFixed(oTG.CreatePoint(0, 0, 0), oTG.CreateUnitVector(1, 0, 0), oTG.CreateUnitVector(0, 1, 0));  //Tao plane1                    
                    PlanarSketch sketch1 = oComDef.Sketches.Add(plane1); //Dung plane1 de tao sketch1
                    SketchPoint centerPoint = (SketchPoint)sketch1.AddByProjectingEntity(axisZ);
                    List<PlanarSketch> arrSketch = new List<PlanarSketch>();
                    arrSketch.Add(sketch1);


                    //Tao vong lap ve cac cap ton


                    int tongsodinh =sola/soLaPerStep;
                    int sochukichan = tongsodinh / sodinhPerPack;
                    int sodinhle = tongsodinh % sodinhPerPack;
                    int vitri = 0;
                   
                    for (int i=0;i<sochukichan;i++)
                    {
                        for(int j =0; j<sodinhPerPack;j++)
                        {                            
                            double lechdinh = j * steplap -  (sodinhPerPack - 1) / 2 * steplap;
                            arrSketch.Add(DrawingLaminate(m_inventorApp, arrSketch[vitri], opartDoc, oComDef, oTG, ccs, hcs, bt, bg, lg, d, lechdinh, soLaPerStep, centerPoint));
                            vitri = vitri + 1;
                        }
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
        public PlanarSketch DrawingLaminate(//Ve mot dinh
                                        Inventor.Application m_inventorApp,
                                        PlanarSketch sketch1,
                                        PartDocument opartDoc,
                                        PartComponentDefinition oComDef,
                                        TransientGeometry oTG,
                                        double ccs,
                                        double hcs,
                                        double bt,
                                        double bg,                                        
                                        double lg,
                                        double d, 
                                        double lechdinh, //Chieu dai lech dinh
                                        double soLaPerStep,
                                        SketchPoint centerPoint
                                                )
        {      
            //===================Xu li bien dau vao truoc khi ve ====================//
            double x0 = 0;
            double y0 = 0;
            lechdinh = lechdinh / 10;
            hcs = hcs / 10;
            ccs = ccs / 10;
            bt = bt / 10;
            bg = bg / 10;
            lg = lg / 10;
            d = d / 10;
            double h_trugiua = hcs + 1 * bg;

            double bedaystep = soLaPerStep * d;
                    
            //---------------------Ve trụ giữa----------//
            SketchPoint trugiua_p1 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - lechdinh, y0 + h_trugiua / 2), false);        //Đỉnh trên
            SketchPoint trugiua_p2 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + lechdinh, y0 - h_trugiua / 2), false);        //Đỉnh dưới
            SketchPoint trugiua_p3 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - bt / 2, y0 + h_trugiua / 2 - bt / 2 + lechdinh), false); //Đỉnh trên bên trái
            SketchPoint trugiua_p4 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + bt / 2, y0 + h_trugiua / 2 - bt / 2 - lechdinh), false); //Đỉnh trên bên phải
            SketchPoint trugiua_p5 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - bt / 2, y0 - h_trugiua / 2 + bt / 2 + lechdinh), false); //Đỉnh dưới bên trái
            SketchPoint trugiua_p6 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + bt / 2, y0 - h_trugiua / 2 + bt / 2 - lechdinh), false); //Đỉnh dưới bên phải
            SketchLine trugiua_ltrentrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p1, trugiua_p3);
            SketchLine trugiua_ltrenphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p1, trugiua_p4);
            SketchLine trugiua_ltrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p3, trugiua_p5);
            SketchLine trugiua_lduoiphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p2, trugiua_p6);
            SketchLine trugiua_lphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p4, trugiua_p6);
            SketchLine trugiua_lduoitrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p5, trugiua_p2);

            sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_ltrentrai, trugiua_ltrenphai, oTG.CreatePoint2d(trugiua_p1.Geometry.X, trugiua_p1.Geometry.Y - 5));
            sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_lduoitrai, trugiua_lduoiphai, oTG.CreatePoint2d(trugiua_p2.Geometry.X, trugiua_p2.Geometry.Y + 5));
            sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_ltrai, trugiua_ltrentrai, oTG.CreatePoint2d(trugiua_p1.Geometry.X - 1, trugiua_p1.Geometry.Y + 10));
            sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_ltrai, trugiua_lduoitrai, oTG.CreatePoint2d(trugiua_p2.Geometry.X - 1, trugiua_p2.Geometry.Y - 1));
            sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p1, trugiua_p3, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(x0, y0));
            sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p2, trugiua_p5, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(x0, y0));
            sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p2, trugiua_p1, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(x0, y0)); //Khoang cach giua hai dinh
            sketch1.GeometricConstraints.AddVertical((SketchEntity)trugiua_ltrai);
            sketch1.GeometricConstraints.AddVertical((SketchEntity)trugiua_lphai);
            sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p3, trugiua_p4, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(x0, y0));

            ObjectCollection oPathSegments1 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
            oPathSegments1.Add(trugiua_lphai);
            oPathSegments1.Add(trugiua_ltrai);
            oPathSegments1.Add(trugiua_ltrenphai);
            oPathSegments1.Add(trugiua_ltrentrai);
            oPathSegments1.Add(trugiua_lduoiphai);
            oPathSegments1.Add(trugiua_lduoitrai);

            Profile oprofile1 = sketch1.Profiles.AddForSolid(false, oPathSegments1);

            //---------------------Ve la gong tren----------//
            SketchPoint p1 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p1.Geometry.X - lg - bg / 2, trugiua_p1.Geometry.Y + bg / 2), false);
            SketchPoint p2 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X + 2 * lg + bg, p1.Geometry.Y), false);
            SketchPoint p3 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p2.Geometry.X - bg, p2.Geometry.Y - bg), false);
            SketchPoint p4 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p3.Geometry.X - lg + bg, p3.Geometry.Y), false);
            SketchPoint p5 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p1.Geometry.X, trugiua_p1.Geometry.Y), false);
            SketchPoint p6 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p5.Geometry.X - bg / 2, p5.Geometry.Y - bg / 2), false);
            SketchPoint p7 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X + bg, p6.Geometry.Y), false);

            SketchLine l12 = sketch1.SketchLines.AddByTwoPoints(p1, p2);
            SketchLine l23 = sketch1.SketchLines.AddByTwoPoints(p2, p3);
            SketchLine l34 = sketch1.SketchLines.AddByTwoPoints(p3, p4);
            SketchLine l45 = sketch1.SketchLines.AddByTwoPoints(p4, p5);
            SketchLine l56 = sketch1.SketchLines.AddByTwoPoints(p5, p6);
            SketchLine l67 = sketch1.SketchLines.AddByTwoPoints(p6, p7);
            SketchLine l71 = sketch1.SketchLines.AddByTwoPoints(p7, p1);

            sketch1.GeometricConstraints.AddCoincident((SketchEntity)l45, (SketchEntity)trugiua_p1);  //Tao rang buoc giua tru giua va gong tren
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)l56, (SketchEntity)trugiua_p1);  //Tao rang buoc giua tru giua va gong tren
            sketch1.GeometricConstraints.AddHorizontal((SketchEntity)l12, true);
            sketch1.GeometricConstraints.AddHorizontal((SketchEntity)l34, true);
            sketch1.GeometricConstraints.AddHorizontal((SketchEntity)l67, true);
            sketch1.DimensionConstraints.AddTwoPointDistance(p1, p5, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y));
            sketch1.DimensionConstraints.AddTwoPointDistance(p1, p5, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y));
            sketch1.DimensionConstraints.AddTwoPointDistance(p2, p5, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y));
            sketch1.DimensionConstraints.AddTwoLineAngle(l45, l56, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y - 5));
            sketch1.DimensionConstraints.AddTwoLineAngle(l34, l45, oTG.CreatePoint2d(p4.Geometry.X + 5, p1.Geometry.Y + 5));
            sketch1.DimensionConstraints.AddTwoPointDistance(p1, p7, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));
            sketch1.DimensionConstraints.AddTwoPointDistance(p1, p7, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));
            sketch1.DimensionConstraints.AddTwoPointDistance(p2, p3, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));
            sketch1.DimensionConstraints.AddTwoPointDistance(p2, p3, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));


            ObjectCollection oPathSegments2 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
            oPathSegments2.Add(l12);
            oPathSegments2.Add(l23);
            oPathSegments2.Add(l34);
            oPathSegments2.Add(l45);
            oPathSegments2.Add(l56);
            oPathSegments2.Add(l67);
            oPathSegments2.Add(l71);
            Profile oprofile2 = sketch1.Profiles.AddForSolid(false, oPathSegments2);

            //-----------------Ve la gong duoi------------//
            SketchPoint gongduoi_p1 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p2.Geometry.X - lg - bg / 2, trugiua_p2.Geometry.Y - bg / 2), false);
            SketchPoint gongduoi_p2 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 2 * lg + bg, gongduoi_p1.Geometry.Y), false);
            SketchPoint gongduoi_p3 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p2.Geometry.X - bg, gongduoi_p2.Geometry.Y + bg), false);
            SketchPoint gongduoi_p4 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p3.Geometry.X - (lg - bg), gongduoi_p3.Geometry.Y), false);
            SketchPoint gongduoi_p5 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p2.Geometry.X, trugiua_p2.Geometry.Y), false);
            SketchPoint gongduoi_p6 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p5.Geometry.X - bg / 2, gongduoi_p5.Geometry.Y + bg / 2), false);
            SketchPoint gongduoi_p7 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X + bg, gongduoi_p6.Geometry.Y), false);

            SketchLine gongduoi_l12 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p1, gongduoi_p2);
            SketchLine gongduoi_l23 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p2, gongduoi_p3);
            SketchLine gongduoi_l34 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p3, gongduoi_p4);
            SketchLine gongduoi_l45 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p4, gongduoi_p5);
            SketchLine gongduoi_l56 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p5, gongduoi_p6);
            SketchLine gongduoi_l67 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p6, gongduoi_p7);
            SketchLine gongduoi_l71 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p7, gongduoi_p1);


            sketch1.GeometricConstraints.AddCoincident((SketchEntity)gongduoi_l45, (SketchEntity)trugiua_p2);  //Tao rang buoc giua tru giua va gong tren
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)gongduoi_l56, (SketchEntity)trugiua_p2);  //Tao rang buoc giua tru giua va gong duoi
            sketch1.GeometricConstraints.AddCollinear((SketchEntity)gongduoi_l45, (SketchEntity)trugiua_lduoiphai);
            sketch1.GeometricConstraints.AddCollinear((SketchEntity)gongduoi_l56, (SketchEntity)trugiua_lduoitrai);
            sketch1.GeometricConstraints.AddCollinear((SketchEntity)gongduoi_l34, (SketchEntity)gongduoi_l67);
            sketch1.GeometricConstraints.AddHorizontal((SketchEntity)gongduoi_l34);
            sketch1.GeometricConstraints.AddHorizontal((SketchEntity)gongduoi_l12);
            sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p1, gongduoi_p7, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
            sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p1, gongduoi_p7, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
            sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p5, gongduoi_p7, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
            sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p5, gongduoi_p7, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
            sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p1, gongduoi_p2, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
            sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p4, gongduoi_p3, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));

            ObjectCollection oPathSegments3 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
            oPathSegments3.Add(gongduoi_l12);
            oPathSegments3.Add(gongduoi_l23);
            oPathSegments3.Add(gongduoi_l34);
            oPathSegments3.Add(gongduoi_l45);
            oPathSegments3.Add(gongduoi_l56);
            oPathSegments3.Add(gongduoi_l67);
            oPathSegments3.Add(gongduoi_l71);
            Profile oprofile3 = sketch1.Profiles.AddForSolid(false, oPathSegments3);
            

            //--------------Ve tru ben trai----------------//
            SketchPoint trutrai_ptrenphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X, p1.Geometry.Y), false);
            SketchPoint trutrai_pduoiphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X, gongduoi_p1.Geometry.Y), false);
            SketchLine trutrai_lphai = sketch1.SketchLines.AddByTwoPoints(trutrai_ptrenphai, trutrai_pduoiphai);
            sketch1.GeometricConstraints.AddVertical((SketchEntity)trutrai_lphai); //vuong goc voi truc toa do
            sketch1.DimensionConstraints.AddTwoPointDistance(trutrai_lphai.EndSketchPoint, trugiua_ltrai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = ccs - bt;
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)trutrai_ptrenphai, (SketchEntity)l71);
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)trutrai_pduoiphai, (SketchEntity)gongduoi_l71);

            SketchPoint trutrai_ptrentrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X, p1.Geometry.Y), false);
            SketchPoint trutrai_pduoitrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X, gongduoi_p1.Geometry.Y), false);
            SketchLine trutrai_ltrai = sketch1.SketchLines.AddByTwoPoints(trutrai_ptrentrai, trutrai_pduoitrai);
            sketch1.GeometricConstraints.AddVertical((SketchEntity)trutrai_ltrai); //vuong goc voi truc toa do
            sketch1.DimensionConstraints.AddTwoPointDistance(trutrai_ltrai.EndSketchPoint, trugiua_ltrai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = ccs;
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)l71, (SketchEntity)trutrai_ptrentrai);
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)trutrai_pduoitrai, (SketchEntity)gongduoi_l71);

            SketchLine trutrai_ltren = sketch1.SketchLines.AddByTwoPoints(trutrai_ptrentrai, trutrai_ptrenphai);
            SketchLine trutrai_lduoi = sketch1.SketchLines.AddByTwoPoints(trutrai_pduoitrai, trutrai_pduoiphai);
           
            ObjectCollection oPathSegments4 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
            oPathSegments4.Add(trutrai_ltrai);
            oPathSegments4.Add(trutrai_lphai);
            oPathSegments4.Add(trutrai_ltren);
            oPathSegments4.Add(trutrai_lduoi);
           
            Profile oprofile4 = sketch1.Profiles.AddForSolid(false, oPathSegments4);

            //-----------Ve tru ben phai-----------------//
            SketchPoint truphai_ptrenphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p2.Geometry.X, p2.Geometry.Y), false);
            SketchPoint truphai_pduoiphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p2.Geometry.X, gongduoi_p2.Geometry.Y), false);
            SketchLine truphai_lphai = sketch1.SketchLines.AddByTwoPoints(truphai_ptrenphai, truphai_pduoiphai);
            sketch1.GeometricConstraints.AddVertical((SketchEntity)truphai_lphai); //vuong goc voi truc toa do
            sketch1.DimensionConstraints.AddTwoPointDistance(truphai_lphai.EndSketchPoint, trugiua_lphai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = ccs;
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)truphai_ptrenphai, (SketchEntity)l23);
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)truphai_pduoiphai, (SketchEntity)gongduoi_l23);

            SketchPoint truphai_ptrentrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p2.Geometry.X, p2.Geometry.Y), false);
            SketchPoint truphai_pduoitrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p2.Geometry.X, gongduoi_p2.Geometry.Y), false);
            SketchLine truphai_ltrai = sketch1.SketchLines.AddByTwoPoints(truphai_ptrentrai, truphai_pduoitrai);
            sketch1.GeometricConstraints.AddVertical((SketchEntity)truphai_ltrai); //vuong goc voi truc toa do
            sketch1.DimensionConstraints.AddTwoPointDistance(truphai_ltrai.EndSketchPoint, trugiua_lphai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p2.Geometry.X + 1, gongduoi_p2.Geometry.Y + 1)).Parameter.Value = ccs - bt;
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)l23, (SketchEntity)truphai_ptrentrai);
            sketch1.GeometricConstraints.AddCoincident((SketchEntity)truphai_pduoitrai, (SketchEntity)gongduoi_l23);

            sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_lphai.EndSketchPoint, centerPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = bt / 2;
            sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p1, centerPoint, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = h_trugiua / 2;

            SketchLine truphai_ltren = sketch1.SketchLines.AddByTwoPoints(truphai_ptrentrai, truphai_ptrenphai);
            SketchLine truphai_lduoi = sketch1.SketchLines.AddByTwoPoints(truphai_pduoitrai, truphai_pduoiphai);


            ObjectCollection oPathSegments5 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
            oPathSegments5.Add(truphai_ltrai);
            oPathSegments5.Add(truphai_lphai);
            oPathSegments5.Add(truphai_ltren);
            oPathSegments5.Add(truphai_lduoi);
            Profile oprofile5 = sketch1.Profiles.AddForSolid(false, oPathSegments5);
            
            m_inventorApp.ActiveDocument.Update();

            ExtrudeFeature oheadExt1 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile1, bedaystep, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
            ExtrudeFeature oheadExt2 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile2, bedaystep, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
            ExtrudeFeature oheadExt3 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile3, bedaystep, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
            ExtrudeFeature oheadExt4 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile4, bedaystep, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
            ExtrudeFeature oheadExt5 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile5, bedaystep, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);

            PlanarSketch reSketch = oComDef.Sketches.Add(oheadExt1.Faces[5]);//tra ve 1 sketch de ve mat tiep theo.
            return reSketch;
        }

        public void DrawingYokeLaminate(double c_cs,double bg,double bt, double lg, double h_trugiua, double x0,double y0)
        {
            bg = bg / 10;
            bt = bt / 10;
            lg = lg / 10;
            c_cs = c_cs / 10;
            h_trugiua = h_trugiua / 10;
            x0 = x0 / 10;
            y0 = y0 / 10;
            int sodinh = 5;
            double lechdinh = 1;

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
                //If not active, create a new inventor session
                if (m_inventorApp != null) //Inventor da duoc mo
                {
                    //Create new part document
                    PartDocument opartDoc = (PartDocument)m_inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, m_inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject, SystemOfMeasureEnum.kMetricSystemOfMeasure));
                    PartComponentDefinition oComDef = (PartComponentDefinition)opartDoc.ComponentDefinition;
                    TransientGeometry oTG = m_inventorApp.TransientGeometry;

                    //Khai bao cac doi tuong can thiet//
                    SketchPoint oCenterPoint;
                    SketchCircle ocircle;
                    DiameterDimConstraint odimConstrain;
                    SketchPoint[] opointArray;
                    SketchLine[] oedgeArray;

                    //Khai bao cac bien nguoi dung
                    oComDef.Parameters.UserParameters.AddByExpression("Ci", "1480", UnitsTypeEnum.kMillimeterLengthUnits);
                    oComDef.Parameters.UserParameters.AddByExpression("Bti", "240", UnitsTypeEnum.kMillimeterLengthUnits);
                    oComDef.Parameters.UserParameters.AddByExpression("Bgi", "340", UnitsTypeEnum.kMillimeterLengthUnits);
                    oComDef.Parameters.UserParameters.AddByExpression("Hi", "340", UnitsTypeEnum.kMillimeterLengthUnits);
                    //Khoi tao goc toa do
                    Inventor.Point oPoint = oTG.CreatePoint(0, 0, 0);
                    //Khoi tao vec to don vi z;
                    Inventor.UnitVector zVector = oTG.CreateUnitVector(0, 0, 1);

                    //Tai tryc axisZ theo phuong Z  
                    WorkAxis axisZ = oComDef.WorkAxes.AddFixed(oPoint, zVector);
                    axisZ.Name = "Truc Z";
                    axisZ.Visible = false;

                    //Tao plabe1
                    WorkPlane plane1 = oComDef.WorkPlanes.AddFixed(oTG.CreatePoint(0, 0, 0), oTG.CreateUnitVector(1, 0, 0), oTG.CreateUnitVector(0, 1, 0));

                    //Dung plane1 de tao sketch1
                    PlanarSketch sketch1 = oComDef.Sketches.Add(plane1);                    
                    SketchPoint centerPoint = (SketchPoint) sketch1.AddByProjectingEntity(axisZ);
                                    
                    //---------------------Ve trụ giữa----------//
                    SketchPoint trugiua_p1 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0-lechdinh, y0+h_trugiua/2),false);        //Đỉnh trên
                    SketchPoint trugiua_p2 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + lechdinh, y0 - h_trugiua / 2), false); //Đỉnh dưới
                    SketchPoint trugiua_p3 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - bt/2, y0 + h_trugiua / 2 - bt/2+lechdinh), false); //Đỉnh trên bên trái
                    SketchPoint trugiua_p4 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + bt / 2, y0 + h_trugiua / 2 - bt / 2 - lechdinh), false); //Đỉnh trên bên phải
                    SketchPoint trugiua_p5 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 - bt / 2, y0 - h_trugiua / 2 + bt / 2 + lechdinh), false); //Đỉnh dưới bên trái
                    SketchPoint trugiua_p6 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(x0 + bt / 2, y0 - h_trugiua / 2 + bt / 2 - lechdinh), false); //Đỉnh dưới bên phải
                    SketchLine trugiua_ltrentrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p1, trugiua_p3);
                    SketchLine trugiua_ltrenphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p1, trugiua_p4);
                    SketchLine trugiua_ltrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p3, trugiua_p5);
                    SketchLine trugiua_lduoiphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p2, trugiua_p6);
                    SketchLine trugiua_lphai = sketch1.SketchLines.AddByTwoPoints(trugiua_p4, trugiua_p6);
                    SketchLine trugiua_lduoitrai = sketch1.SketchLines.AddByTwoPoints(trugiua_p5, trugiua_p2);
                    
                    sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_ltrentrai, trugiua_ltrenphai, oTG.CreatePoint2d(trugiua_p1.Geometry.X, trugiua_p1.Geometry.Y - 5));
                    sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_lduoitrai, trugiua_lduoiphai, oTG.CreatePoint2d(trugiua_p2.Geometry.X, trugiua_p2.Geometry.Y + 5));
                    sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_ltrai, trugiua_ltrentrai, oTG.CreatePoint2d(trugiua_p1.Geometry.X-1, trugiua_p1.Geometry.Y + 10));                    
                    sketch1.DimensionConstraints.AddTwoLineAngle(trugiua_ltrai, trugiua_lduoitrai, oTG.CreatePoint2d(trugiua_p2.Geometry.X - 1, trugiua_p2.Geometry.Y -1));
                    sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p1, trugiua_p3,DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(x0, y0));
                    sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p2, trugiua_p5, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(x0, y0));
                    sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p2, trugiua_p1, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(x0, y0)); //Khoang cach giua hai dinh
                    sketch1.GeometricConstraints.AddVertical((SketchEntity)trugiua_ltrai);
                    sketch1.GeometricConstraints.AddVertical((SketchEntity)trugiua_lphai);
                    sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p3, trugiua_p4, DimensionOrientationEnum.kHorizontalDim,oTG.CreatePoint2d(x0,y0));

                    ObjectCollection oPathSegments1 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
                    oPathSegments1.Add(trugiua_lphai);
                    oPathSegments1.Add(trugiua_ltrai);
                    oPathSegments1.Add(trugiua_ltrenphai);
                    oPathSegments1.Add(trugiua_ltrentrai);
                    oPathSegments1.Add(trugiua_lduoiphai);
                    oPathSegments1.Add(trugiua_lduoitrai);

                    Profile oprofile1 = sketch1.Profiles.AddForSolid(false,oPathSegments1);                    

                    //---------------------Ve la gong tren----------//
                    SketchPoint p1 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p1.Geometry.X-lg-bg/2, trugiua_p1.Geometry.Y+ bg/2), false);
                    SketchPoint p2 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X+2*lg+bg, p1.Geometry.Y), false);
                    SketchPoint p3 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p2.Geometry.X-bg, p2.Geometry.Y-bg), false);
                    SketchPoint p4 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p3.Geometry.X-lg+bg, p3.Geometry.Y), false);
                    SketchPoint p5 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p1.Geometry.X, trugiua_p1.Geometry.Y), false);
                    SketchPoint p6 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p5.Geometry.X-bg/2, p5.Geometry.Y-bg/2), false);
                    SketchPoint p7 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X+bg,p6.Geometry.Y), false);
                    
                    SketchLine l12 = sketch1.SketchLines.AddByTwoPoints(p1, p2);
                    SketchLine l23 = sketch1.SketchLines.AddByTwoPoints(p2, p3);
                    SketchLine l34 = sketch1.SketchLines.AddByTwoPoints(p3, p4);
                    SketchLine l45 = sketch1.SketchLines.AddByTwoPoints(p4, p5);
                    SketchLine l56 = sketch1.SketchLines.AddByTwoPoints(p5, p6);
                    SketchLine l67 = sketch1.SketchLines.AddByTwoPoints(p6, p7);
                    SketchLine l71 = sketch1.SketchLines.AddByTwoPoints(p7, p1);

                    sketch1.GeometricConstraints.AddHorizontal((SketchEntity)l12, true);
                    sketch1.GeometricConstraints.AddHorizontal((SketchEntity)l34, true);
                    sketch1.GeometricConstraints.AddHorizontal((SketchEntity)l67, true);

                    sketch1.DimensionConstraints.AddTwoPointDistance(p1, p5, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y));
                    sketch1.DimensionConstraints.AddTwoPointDistance(p1, p5, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y));
                    sketch1.DimensionConstraints.AddTwoPointDistance(p2, p5, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y));

                    sketch1.DimensionConstraints.AddTwoLineAngle(l45, l56, oTG.CreatePoint2d(p5.Geometry.X, p5.Geometry.Y - 5));
                    sketch1.DimensionConstraints.AddTwoLineAngle(l34, l45, oTG.CreatePoint2d(p4.Geometry.X + 5, p1.Geometry.Y + 5));

                    sketch1.DimensionConstraints.AddTwoPointDistance(p1, p7,DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));
                    sketch1.DimensionConstraints.AddTwoPointDistance(p1, p7, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));
                    sketch1.DimensionConstraints.AddTwoPointDistance(p2, p3, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));
                    sketch1.DimensionConstraints.AddTwoPointDistance(p2, p3, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(p1.Geometry.X + 5, p1.Geometry.Y - 2));

                    ObjectCollection oPathSegments2 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
                    oPathSegments2.Add(l12);
                    oPathSegments2.Add(l23);
                    oPathSegments2.Add(l34);
                    oPathSegments2.Add(l45);
                    oPathSegments2.Add(l56);
                    oPathSegments2.Add(l67);
                    oPathSegments2.Add(l71);
                    Profile oprofile2 = sketch1.Profiles.AddForSolid(false, oPathSegments2);
                    
                    //-----------------Ve la gong duoi------------//
                    SketchPoint gongduoi_p1 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p2.Geometry.X - lg - bg / 2, trugiua_p2.Geometry.Y - bg / 2), false);
                    SketchPoint gongduoi_p2 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 2*lg + bg, gongduoi_p1.Geometry.Y), false);
                    SketchPoint gongduoi_p3 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p2.Geometry.X - bg, gongduoi_p2.Geometry.Y + bg), false);
                    SketchPoint gongduoi_p4 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p3.Geometry.X - (lg - bg), gongduoi_p3.Geometry.Y), false);
                    SketchPoint gongduoi_p5 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(trugiua_p2.Geometry.X, trugiua_p2.Geometry.Y), false);
                    SketchPoint gongduoi_p6 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p5.Geometry.X - bg / 2, gongduoi_p5.Geometry.Y + bg / 2), false);
                    SketchPoint gongduoi_p7 = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X + bg, gongduoi_p6.Geometry.Y), false);

                    SketchLine gongduoi_l12 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p1, gongduoi_p2);
                    SketchLine gongduoi_l23 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p2, gongduoi_p3);
                    SketchLine gongduoi_l34 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p3, gongduoi_p4);
                    SketchLine gongduoi_l45 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p4, gongduoi_p5);
                    SketchLine gongduoi_l56 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p5, gongduoi_p6);
                    SketchLine gongduoi_l67 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p6, gongduoi_p7);
                    SketchLine gongduoi_l71 = sketch1.SketchLines.AddByTwoPoints(gongduoi_p7, gongduoi_p1);

                    sketch1.GeometricConstraints.AddCollinear((SketchEntity)gongduoi_l34, (SketchEntity)gongduoi_l67);
                    sketch1.GeometricConstraints.AddHorizontal((SketchEntity)gongduoi_l34);
                    sketch1.GeometricConstraints.AddHorizontal((SketchEntity)gongduoi_l12);
                    sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p1, gongduoi_p7, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X+1,gongduoi_p1.Geometry.Y+1));
                    sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p1, gongduoi_p7, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
                    sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p5, gongduoi_p7, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
                    sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p5, gongduoi_p7, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
                    sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p1, gongduoi_p2, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));
                    sketch1.DimensionConstraints.AddTwoPointDistance(gongduoi_p4, gongduoi_p3, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1));

                    ObjectCollection oPathSegments3 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
                    oPathSegments3.Add(gongduoi_l12);
                    oPathSegments3.Add(gongduoi_l23);
                    oPathSegments3.Add(gongduoi_l34);
                    oPathSegments3.Add(gongduoi_l45);
                    oPathSegments3.Add(gongduoi_l56);
                    oPathSegments3.Add(gongduoi_l67);
                    oPathSegments3.Add(gongduoi_l71);
                    Profile oprofile3 = sketch1.Profiles.AddForSolid(false, oPathSegments3);

                    //--------------Ve tru ben trai----------------//
                    SketchPoint trutrai_ptrenphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X , p1.Geometry.Y ), false);
                    SketchPoint trutrai_pduoiphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X, gongduoi_p1.Geometry.Y), false);
                    SketchLine trutrai_lphai = sketch1.SketchLines.AddByTwoPoints(trutrai_ptrenphai, trutrai_pduoiphai);
                    sketch1.GeometricConstraints.AddVertical((SketchEntity)trutrai_lphai); //vuong goc voi truc toa do
                    sketch1.DimensionConstraints.AddTwoPointDistance(trutrai_lphai.EndSketchPoint, trugiua_ltrai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value=c_cs-bt;
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)trutrai_ptrenphai, (SketchEntity)l71);
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)trutrai_pduoiphai, (SketchEntity)gongduoi_l71);

                    SketchPoint trutrai_ptrentrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p1.Geometry.X, p1.Geometry.Y), false);
                    SketchPoint trutrai_pduoitrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p1.Geometry.X, gongduoi_p1.Geometry.Y), false);
                    SketchLine trutrai_ltrai= sketch1.SketchLines.AddByTwoPoints(trutrai_ptrentrai, trutrai_pduoitrai);
                    sketch1.GeometricConstraints.AddVertical((SketchEntity)trutrai_ltrai); //vuong goc voi truc toa do
                    sketch1.DimensionConstraints.AddTwoPointDistance(trutrai_ltrai.EndSketchPoint, trugiua_ltrai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = c_cs ;
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)l71,(SketchEntity)trutrai_ptrentrai );
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)trutrai_pduoitrai, (SketchEntity)gongduoi_l71);

                    SketchLine trutrai_ltren = sketch1.SketchLines.AddByTwoPoints(trutrai_ptrentrai, trutrai_ptrenphai);
                    SketchLine trutrai_lduoi = sketch1.SketchLines.AddByTwoPoints(trutrai_pduoitrai, trutrai_pduoiphai);

                    ObjectCollection oPathSegments4 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
                    oPathSegments4.Add(trutrai_ltrai);
                    oPathSegments4.Add(trutrai_lphai);
                    oPathSegments4.Add(trutrai_ltren);
                    oPathSegments4.Add(trutrai_lduoi);                    
                    Profile oprofile4= sketch1.Profiles.AddForSolid(false, oPathSegments4);

                    //-----------Ve tru ben phai-----------------//
   

                    SketchPoint truphai_ptrenphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p2.Geometry.X, p2.Geometry.Y), false);
                    SketchPoint truphai_pduoiphai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p2.Geometry.X, gongduoi_p2.Geometry.Y), false);
                    SketchLine truphai_lphai = sketch1.SketchLines.AddByTwoPoints(truphai_ptrenphai, truphai_pduoiphai);
                    sketch1.GeometricConstraints.AddVertical((SketchEntity)truphai_lphai); //vuong goc voi truc toa do
                    sketch1.DimensionConstraints.AddTwoPointDistance(truphai_lphai.EndSketchPoint, trugiua_lphai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = c_cs;
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)truphai_ptrenphai, (SketchEntity)l23);
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)truphai_pduoiphai, (SketchEntity)gongduoi_l23);

                    SketchPoint truphai_ptrentrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(p2.Geometry.X, p2.Geometry.Y), false);
                    SketchPoint truphai_pduoitrai = sketch1.SketchPoints.Add(oTG.CreatePoint2d(gongduoi_p2.Geometry.X, gongduoi_p2.Geometry.Y), false);
                    SketchLine truphai_ltrai = sketch1.SketchLines.AddByTwoPoints(truphai_ptrentrai, truphai_pduoitrai);
                    sketch1.GeometricConstraints.AddVertical((SketchEntity)truphai_ltrai); //vuong goc voi truc toa do
                    sketch1.DimensionConstraints.AddTwoPointDistance(truphai_ltrai.EndSketchPoint, trugiua_lphai.StartSketchPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p2.Geometry.X + 1, gongduoi_p2.Geometry.Y + 1)).Parameter.Value = c_cs - bt;
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)l23, (SketchEntity)truphai_ptrentrai);
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)truphai_pduoitrai, (SketchEntity)gongduoi_l23);

                    SketchLine truphai_ltren = sketch1.SketchLines.AddByTwoPoints(truphai_ptrentrai, truphai_ptrenphai);
                    SketchLine truphai_lduoi = sketch1.SketchLines.AddByTwoPoints(truphai_pduoitrai, truphai_pduoiphai);

                    ObjectCollection oPathSegments5 = m_inventorApp.TransientObjects.CreateObjectCollection(); //Tao collection chua cac doi tuong
                    oPathSegments5.Add(truphai_ltrai);
                    oPathSegments5.Add(truphai_lphai);
                    oPathSegments5.Add(truphai_ltren);
                    oPathSegments5.Add(truphai_lduoi);
                    Profile oprofile5 = sketch1.Profiles.AddForSolid(false, oPathSegments5);

                    //Tao cac rang buoc 

                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)l45, (SketchEntity)trugiua_p1);  //Tao rang buoc giua tru giua va gong tren
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)l56, (SketchEntity)trugiua_p1);  //Tao rang buoc giua tru giua va gong tren

                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)gongduoi_l45, (SketchEntity)trugiua_p2);  //Tao rang buoc giua tru giua va gong tren
                    sketch1.GeometricConstraints.AddCoincident((SketchEntity)gongduoi_l56, (SketchEntity)trugiua_p2);  //Tao rang buoc giua tru giua va gong duoi

                    sketch1.GeometricConstraints.AddCollinear((SketchEntity)gongduoi_l45, (SketchEntity)trugiua_lduoiphai);
                    sketch1.GeometricConstraints.AddCollinear((SketchEntity)gongduoi_l56, (SketchEntity)trugiua_lduoitrai);

                    sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_lphai.EndSketchPoint, centerPoint, DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = bt/2;
                    sketch1.DimensionConstraints.AddTwoPointDistance(trugiua_p1, centerPoint, DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(gongduoi_p1.Geometry.X + 1, gongduoi_p1.Geometry.Y + 1)).Parameter.Value = h_trugiua / 2;
                    
                    m_inventorApp.ActiveDocument.Update();
                             
                    ExtrudeFeature oheadExt1 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile1, 2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
                    ExtrudeFeature oheadExt2 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile2, 2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
                    ExtrudeFeature oheadExt3 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile3, 2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
                    ExtrudeFeature oheadExt4 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile4, 2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);
                    ExtrudeFeature oheadExt5 = oComDef.Features.ExtrudeFeatures.AddByDistanceExtent(oprofile5, 2, PartFeatureExtentDirectionEnum.kPositiveExtentDirection, PartFeatureOperationEnum.kJoinOperation);

                    //odimConstrain.Parameter.Expression = "HeadDiameter";
                    //Extrude doi tuong



                    ////Tao sket moi tren oheadExt
                    //PlanarSketch plane2 = oComDef.Sketches.Add(oheadExt.Faces[2], true);
                }
                else //Inventor chua duoc mo
                {
                    //Inventor chua duoc mo
                    MessageBox.Show("Bạn phải mở inventor lên đã");
                    //Type inventorAppType = Type.GetTypeFromProgID("Inventor.Application");
                    //m_inventorApp = Activator.CreateInstance(inventorAppType) as Inventor.Application;
                    //m_quitInventor = true;
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
