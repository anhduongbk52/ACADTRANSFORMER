using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ACADTRANSFORMER.MBAOBJECT;
using Microsoft.Win32;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using ACADTRANSFORMER.THUVIENCAD;


namespace ACADTRANSFORMER.GIAODIEN
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        MBA1 mba = new MBA1();
        MBAOBJECT.SODOTRAI sodotrai = new SODOTRAI();
        public Window1()
        {
            InitializeComponent();            
            this.DataContext = mba;
            
        }
        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
        private void BtnN_Click(object sender, RoutedEventArgs e)
        {
            mba.Tinhsolaton();
        }
        private void BtnHbg_Click(object sender, RoutedEventArgs e)
        {
            mba.AutoHbg();  
        }
        private void BtnBg_Click(object sender, RoutedEventArgs e)
        {
            mba.AutoBg();
        }
        private void BtnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            DATA.INOUT.ExportDataToExcel(mba);
        }
        private void BtnImportFormExcel_Click(object sender, RoutedEventArgs e)
        {
            string path;
             System.Windows.Forms.OpenFileDialog opFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            opFileDialog1.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls";
           
            if (opFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = opFileDialog1.FileName;
            }
            else path = null;
            try
            {
                MBA1 temp = DATA.INOUT.ImportDataFromExcel(path);
                mba.Bt = temp.Bt;
                mba.Bg = temp.Bg;
                mba.Hbg = temp.Hbg;
                mba.N = temp.N;
                mba.Thongdaucapton = temp.Thongdaucapton;
                mba.Diameter = temp.Diameter;
                mba.Ccs = temp.Ccs;
                mba.Hcs = temp.Hcs;
                mba.Hr = temp.Hr;
                mba.Delta = temp.Delta;
                mba.Ke = temp.Ke;
                mba.Lt = temp.Lt;
                mba.Ks = temp.Ks;
                mba.ChieuEp = temp.ChieuEp;
                mba.Section = temp.Section;
                mba.Uv = temp.Uv;
                mba.Po = temp.Po;
                mba.Poyc = temp.Poyc;

                mba.Winding1.D1t = temp.Winding1.D1t;
                mba.Winding1.D2t = temp.Winding1.D2t;
                mba.Winding1.D3t = temp.Winding1.D3t;
                mba.Winding1.D4t = temp.Winding1.D4t;
                mba.Winding1.D5t = temp.Winding1.D5t;

                mba.Winding1.D1n = temp.Winding1.D1n;
                mba.Winding1.D2n = temp.Winding1.D2n;
                mba.Winding1.D3n = temp.Winding1.D3n;
                mba.Winding1.D4n = temp.Winding1.D4n;
                mba.Winding1.D5n = temp.Winding1.D5n;

                mba.Winding1.H1 = temp.Winding1.H1;
                mba.Winding1.H2 = temp.Winding1.H2;
                mba.Winding1.H3 = temp.Winding1.H3;
                mba.Winding1.H4 = temp.Winding1.H4;
                mba.Winding1.H5 = temp.Winding1.H5;

                mba.Winding1.Demdau1 = temp.Winding1.H1;
                mba.Winding1.Demdau2 = temp.Winding1.H2;
                mba.Winding1.Demdau3 = temp.Winding1.H3;
                mba.Winding1.Demdau4 = temp.Winding1.H4;
                mba.Winding1.Demdau5 = temp.Winding1.H5;

                MessageBox.Show("DONE");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
           
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Xml Files|*.xml";
            saveFileDialog1.Title = "Save an Xml file";
            if (saveFileDialog1.ShowDialog() == true)
            {
                string  path = saveFileDialog1.FileName;
                DATA.INOUT.SaveToFileXml(mba,path);
            }            
        }
        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            string path;
            OpenFileDialog opFileDialog1 = new OpenFileDialog();
            opFileDialog1.Filter = "Xml Files|*.xml";

            if (opFileDialog1.ShowDialog() == true)
            {
                path = opFileDialog1.FileName;
            }
            else path = null;
            try
            {
                MBA1 temp = DATA.INOUT.ReadDataFormXml(path);

                mba.IsGiamChieuDaiRau = temp.IsGiamChieuDaiRau;
                mba.IsMultiMaterial = temp.IsMultiMaterial;
                mba.RealDeltai = temp.RealDeltai;
                mba.Bt = temp.Bt;
                mba.Bg = temp.Bg;
                mba.Hbg = temp.Hbg;
                mba.N = temp.N;
                mba.RealDeltai = temp.RealDeltai;
                mba.Thongdaucapton = temp.Thongdaucapton;
                mba.Diameter = temp.Diameter;
                mba.Ccs = temp.Ccs;
                mba.Hcs = temp.Hcs;
                mba.Hr = temp.Hr;
                mba.Delta = temp.Delta;
                mba.Ke = temp.Ke;
                mba.Lt = temp.Lt;
                mba.Ks = temp.Ks;
                mba.ChieuEp = temp.ChieuEp;
                mba.Section = temp.Section;
                mba.Uv = temp.Uv;
                mba.Po = temp.Po;
                mba.Poyc = temp.Poyc;

                mba.Winding1.D1t = temp.Winding1.D1t;
                mba.Winding1.D2t = temp.Winding1.D2t;
                mba.Winding1.D3t = temp.Winding1.D3t;
                mba.Winding1.D4t = temp.Winding1.D4t;
                mba.Winding1.D5t = temp.Winding1.D5t;

                mba.Winding1.D1n = temp.Winding1.D1n;
                mba.Winding1.D2n = temp.Winding1.D2n;
                mba.Winding1.D3n = temp.Winding1.D3n;
                mba.Winding1.D4n = temp.Winding1.D4n;
                mba.Winding1.D5n = temp.Winding1.D5n;

                mba.Winding1.H1 = temp.Winding1.H1;
                mba.Winding1.H2 = temp.Winding1.H2;
                mba.Winding1.H3 = temp.Winding1.H3;
                mba.Winding1.H4 = temp.Winding1.H4;
                mba.Winding1.H5 = temp.Winding1.H5;

                mba.Winding1.L01 = temp.Winding1.L01;
                mba.Winding1.L02 = temp.Winding1.L02;

                mba.Winding1.Demdau1 = temp.Winding1.Demdau1;
                mba.Winding1.Demdau2 = temp.Winding1.Demdau2;
                mba.Winding1.Demdau3 = temp.Winding1.Demdau3;
                mba.Winding1.Demdau4 = temp.Winding1.Demdau4;
                mba.Winding1.Demdau5 = temp.Winding1.Demdau5;

                mba.Winding1.NCanDoc = temp.Winding1.NCanDoc;

                MessageBox.Show("DONE");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }           
            
        }
        private void BtnVeMatCatTru_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            double[] toadotam = THUVIENCAD.MyAcad.GetPosition();
            if (toadotam[2] == 1)
            {
                mba.vematcattru(toadotam[0],toadotam[1]);
            }
            MessageBoxResult dlg = MessageBox.Show("Bạn có muốn tiếp tục", "Done", MessageBoxButton.YesNo);
            if (dlg == MessageBoxResult.Yes)
            {
                this.Show();
            }
            else this.Close();
        }
        private void BtnVeMatCatGong_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            double[] toadotam = THUVIENCAD.MyAcad.GetPosition();
            if (toadotam[2] == 1)
            {
                mba.vematcatgong(toadotam[0], toadotam[1]);
            }
            MessageBoxResult dlg = MessageBox.Show("Bạn có muốn tiếp tục", "Done", MessageBoxButton.YesNo);
            if (dlg == MessageBoxResult.Yes)
            {
                this.Show();
            }
            else this.Close();
        }
        private void BtnVeHinhChieu_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            double[] toadotam = THUVIENCAD.MyAcad.GetPosition();
            if (toadotam[2] == 1)
            {
                mba.vehinhchieu(toadotam[0], toadotam[1]);
            }

            MessageBoxResult dlg = MessageBox.Show("Bạn có muốn tiếp tục", "Done", MessageBoxButton.YesNo);
            if (dlg == MessageBoxResult.Yes)
            {
                this.Show();
            }
            else this.Close();
        }
        private void BtnVeBoiDay_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            double[] toadotam = THUVIENCAD.MyAcad.GetPosition();
            if (toadotam[2] == 1)
            {
                mba.veboiday(toadotam[0], toadotam[1]);
            }

            MessageBoxResult dlg = MessageBox.Show("Bạn có muốn tiếp tục", "Done", MessageBoxButton.YesNo);
            if (dlg == MessageBoxResult.Yes)
            {
                this.Show();
            }
            else this.Close();
        }
        private void BtnInsertBVMachTu_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            double[] toadogoc = THUVIENCAD.MyAcad.GetPosition();
            if (toadogoc[2] == 1)
            {      
                THUVIENCAD.MyAcad.AddrecWidthHigh(toadogoc[0] + 3978 + 160, toadogoc[1] + 7444.8 - 985 - 210, 6290, 1195); // Khung bao 1
                THUVIENCAD.MyAcad.AddLine2d(toadogoc[0] + 3978 + 160, toadogoc[1] + 7444.8 - 985 - 210 + 945, toadogoc[0] + 3978 + 160 + 6290, toadogoc[1] + 7444.8 - 985 - 210 + 945);
                THUVIENCAD.MyAcad.AddLine2d(toadogoc[0] + 3978 + 160 + 1890, toadogoc[1] + 7444.8 - 985 - 210, toadogoc[0] + 3978 + 160 + 1890, toadogoc[1] + 7444.8 - 985 - 210+945);
                THUVIENCAD.MyAcad.AddLine2d(toadogoc[0] + 3978 + 160 + 1890 + 1650, toadogoc[1] + 7444.8 - 985 - 210, toadogoc[0] + 3978 + 160 + 1890 + 1650, toadogoc[1] + 7444.8 - 985 - 210 + 945);
                
                string mtext0 = "D=0; V=10; S=5; C=0; MODE = LLAP; DIR=MIN; N.SAME=2";
                THUVIENCAD.MyAcad.InserMText(toadogoc[0] + 6450, toadogoc[1] + 7300, 66, mtext0);
                THUVIENCAD.MyAcad.InsertBlockBorderA3(toadogoc[0], toadogoc[1], 26.4);
                THUVIENCAD.MyAcad.InsertBlockTitleA3(toadogoc[0] + 215 * 26.4, toadogoc[1] + 5 * 26.4, 26.4);
                mba.insertTableCore1(toadogoc[0] + 3978 +160 , toadogoc[1] + 7444.8 - 985-210);
                Point2d p1 = new Point2d(toadogoc[0] + 4290, toadogoc[1] + 6710);
                THUVIENCAD.MyAcad.DrawTruGiua(new Point3d(p1.X+300,p1.Y-100, 0), true);
                THUVIENCAD.MyAcad.DrawTruBen(new Point2d(p1.X + 1810+300, p1.Y-215), 185, 865, true);
                THUVIENCAD.MyAcad.DrawLaGong(new Point2d(p1.X + 4000-300, p1.Y-215), 185, 865);
                string txtLine2 = "";
                string txtLine3 = "- Uv = "+mba.Uv+" (vol/ vòng).\\P";
                string txtLine4 = "- Po = "+mba.Po+" kW (Yêu cầu theo HĐ Po = 22kW)\\P";
                for (int i = 0; i < mba.Bt.Length;i++ )
                {
                    if (mba.Thongdaucapton[0] != 0) txtLine2 = "- Ở chính giữa cấp A = " + mba.Bt[0] + " có rãnh thông dầu 4 mm được  dán trực tiếp lên lá thép.\\P";
                    else if (mba.Thongdaucapton[i] != 0) txtLine2 = "- Ở giữa cấp A = " + mba.Bt[i - 1] + " và cấp A = " + mba.Bt[i] + " có rãnh thông dầu 4 mm được  dán trực tiếp lên lá thép.\\P";       
                }
                string mtext = " YÊU CẦU KỸ THUẬT:\\P" +
                                    "- Chế tạo theo \" QUY TRÌNH CHẾ TẠO MẠCH TỪ MÁY BIẾN ÁP \"  mã số : CN-QT-75-01, ban hành tháng 05 năm 2013.\\P" +
                                    txtLine2+
                                    txtLine3+
                                    txtLine4+
                                    "- io ≤ 0,1%.\\P";
                THUVIENCAD.MyAcad.InserMText(toadogoc[0] + 500, toadogoc[1] + 1000, 66, mtext);

                THUVIENCAD.MyAcad.DrawMinhHoaGepTon1(new Point2d(p1.X + 1200, p1.Y + 5750), 1);
            }
            MessageBoxResult dlg = MessageBox.Show("Bạn có muốn tiếp tục", "Done", MessageBoxButton.YesNo);
            if (dlg == MessageBoxResult.Yes)
            {
                this.Show();
            }
            else this.Close();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //-------------------Tao style chuan-----------------//
            THUVIENCAD.StyleAutocad.LoadLineType("Continuous");
            THUVIENCAD.StyleAutocad.LoadLineType("DASHED");
            THUVIENCAD.StyleAutocad.LoadLineType("CENTER2");
            THUVIENCAD.StyleAutocad.LoadLineType("DASHDOT");
            
            //1. Khoi tao cac layer chuan
            THUVIENCAD.StyleAutocad.CreatLayer("Visible (ISO)", "Continuous", 7, LineWeight.LineWeight030, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Visible Narrow (ISO)", "Continuous", 4, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Title (ISO)", "Continuous", 7, LineWeight.LineWeight030, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Symbol (ISO)", "Continuous", 7, LineWeight.LineWeight030, true);
            THUVIENCAD.StyleAutocad.CreatLayer("NoPlot", "Continuous", 8, LineWeight.ByLineWeightDefault, false);
            THUVIENCAD.StyleAutocad.CreatLayer("Hidden (ISO)", "DASHED", 4, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Hatch (ISO)", "Continuous", 8, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Dimension (ISO)", "Continuous", 3, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Center Mark (ISO)", "CENTER2", 2, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Centerline (ISO)", "DASHDOT", 2, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Border (ISO)", "Continuous", 7, LineWeight.LineWeight035, true);
            
            //Khởi tạo text style
            THUVIENCAD.MyAcad.CreatTextStyle("STANDARD_DUONG", "TCVN 7284.ttf",0);
            THUVIENCAD.MyAcad.CreatTextStyle("X1_BTK", "TCVN 7284.ttf", 3);
            THUVIENCAD.MyAcad.CreatTextStyle("X2_BTK", "TCVN 7284.ttf", 5);
            THUVIENCAD.MyAcad.CreatTextStyle("X3_BTK", "TCVN 7284.ttf", 7.5);
            THUVIENCAD.MyAcad.CreatTextStyle("X4_BTK", "TCVN 7284.ttf", 10);
            THUVIENCAD.MyAcad.CreatTextStyle("X5_BTK", "TCVN 7284.ttf", 12.5);
            THUVIENCAD.MyAcad.CreatTextStyle("X6_BTK", "TCVN 7284.ttf", 15);
            THUVIENCAD.MyAcad.CreatTextStyle("X8_BTK", "TCVN 7284.ttf", 24);
            THUVIENCAD.MyAcad.CreatTextStyle("X10_BTK", "TCVN 7284.ttf", 25);
            THUVIENCAD.MyAcad.CreatTextStyle("X15_BTK", "TCVN 7284.ttf", 37.5);
            THUVIENCAD.MyAcad.CreatTextStyle("X20_BTK", "TCVN 7284.ttf", 50);
            THUVIENCAD.MyAcad.CreatTextStyle("X26.4_BTK","TCVN 7284.ttf", 66);
            THUVIENCAD.MyAcad.CreatTextStyle("X36_BTK", "TCVN 7284.ttf", 90);
            THUVIENCAD.MyAcad.CreatTextStyle("X40_BTK", "TCVN 7284.ttf", 100);
            
            THUVIENCAD.MyAcad.CreatCostumProperties("Vẽ", "Đỗ Ánh Dương");
            THUVIENCAD.MyAcad.CreatCostumProperties("Thiết kế", "Đỗ Ánh Dương");
            THUVIENCAD.MyAcad.CreatCostumProperties("Kiểm soát", "Lê Hải Quân");
            THUVIENCAD.MyAcad.CreatCostumProperties("Ban thiết kế", "Ng.Duy Linh");
            THUVIENCAD.MyAcad.CreatCostumProperties("Duyệt", "Hồ Đức Thanh");
            THUVIENCAD.MyAcad.CreatCostumProperties("DateTime", "01-03-2016");
            THUVIENCAD.MyAcad.CreatCostumProperties("Công suất", "63M");
            THUVIENCAD.MyAcad.CreatCostumProperties("Điện áp", "115/38,5/23");
            THUVIENCAD.MyAcad.CreatCostumProperties("Trạm", "--");
            THUVIENCAD.MyAcad.CreatCostumProperties("MS", "174735-xx");
            //Khởi tạo các dimension style
            THUVIENCAD.MyAcad.CreatDimStyle("X1_BTK", "X1_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X2_BTK", "X2_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X3_BTK", "X3_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X4_BTK", "X4_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X5_BTK", "X5_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X6_BTK", "X6_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X8_BTK", "X8_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X10_BTK", "X10_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X15_BTK", "X15_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X20_BTK", "X20_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X26.4_BTK","X26.4_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X36_BTK", "X36_BTK");
            THUVIENCAD.MyAcad.CreatDimStyle("X40_BTK", "X40_BTK"); 
        }
        private void ReDraw_Click(object sender, RoutedEventArgs e)
        {
                      
            double x0 = khungve.Width/2; // toa do tam
            double y0 = khungve.Width/2; // toa do tam

            System.Windows.Shapes.Ellipse cirDuongkinh; //Đường tròn bao trụ 
            khungve.Children.Clear();
            cirDuongkinh = new System.Windows.Shapes.Ellipse 
             {
                 Stroke = Brushes.Black,
                 StrokeThickness = 1,
                 Height = khungve.Width ,
                 Width= khungve.Width,
                 StrokeDashArray = new System.Windows.Media.DoubleCollection() { 10 }
             };
             Canvas.SetTop(cirDuongkinh, 0);
             Canvas.SetLeft(cirDuongkinh, 0);
             khungve.Children.Add(cirDuongkinh);
           
            //Ve duong tam truc ngang
             System.Windows.Shapes.Line line1 = new System.Windows.Shapes.Line
             {
                 Stroke = Brushes.LightBlue,
                 StrokeThickness = 1,
                 StrokeDashArray = new System.Windows.Media.DoubleCollection() { 10, 2 , 10 },
                 X1 = 0,               Y1 = khungve.Width / 2,
                 X2 = khungve.Width,   Y2 = khungve.Width/2,
             };
             Canvas.SetTop(line1, 0);
             Canvas.SetLeft(line1, 0);            
             khungve.Children.Add(line1);

             //Ve duong tam truc dung
             System.Windows.Shapes.Line line2 = new System.Windows.Shapes.Line
             {
                 Stroke = Brushes.LightBlue,
                 StrokeThickness = 1,
                 StrokeDashArray = new System.Windows.Media.DoubleCollection() { 10, 2, 10 },
                 X1 = khungve.Width/2,
                 Y1 = 0,
                 X2 = khungve.Width/2,
                 Y2 = khungve.Width
                 
             };
             Canvas.SetTop(line2, 0);
             Canvas.SetLeft(line2, 0);
             khungve.Children.Add(line2);
             Random r = new Random();
             double factor = 0;
             if(mba.Diameter!=0)  factor = khungve.Width / mba.Diameter;             

             int vol = mba.Bt.Length;
             double[] _offsetHB = new double[vol];
            
            //Xac dinh toa do tam cua cap ton
            _offsetHB[0] = 0;
             for (int i = 1; i < vol; i++)
             {
                 if (mba.Bg[i] != 0) _offsetHB[i] = -(mba.Bg[0] - mba.Bg[i]) / 2 + mba.Hbg[i];
             }
            //Ve cac cap ton
             double j = x0 + (double)mba.Thongdaucapton[0] * mba.Hr * factor/2;

            
             for (int i = 0; i < vol; i++)
             {

                 Color mycolor = Colors.AliceBlue;
                 if (mba.Bg[i] != 0)
                 {
                     GrapWPF.drawRecCenterLeft(j + ((double)mba.Thongdaucapton[i] * mba.Hr) * factor, y0 + _offsetHB[i] * factor, (mba.Beday[i] / 2) * factor, mba.Bg[i] * factor, 0.5, mycolor, khungve, "Bg=" + mba.Bg[i].ToString() + "; Hbg=" + mba.Hbg[i].ToString());
                     GrapWPF.drawRecCenterRight(2 * x0 - (j + ((double)mba.Thongdaucapton[i] * mba.Hr) * factor), y0 + _offsetHB[i] * factor, (mba.Beday[i] / 2) * factor, mba.Bg[i] * factor, 0.5, mycolor, khungve, "Bg=" + mba.Bg[i].ToString() + "; Hbg=" + mba.Hbg[i].ToString());
                     j = j + (mba.Beday[i] / 2 + (double)mba.Thongdaucapton[i] * mba.Hr) * factor;
                 }              
             }            
        }
        private void MenuIventorInsertCore_Click(object sender, RoutedEventArgs e)
        {
            INVENTORFUNC.INVFUNC inventor1 = new INVENTORFUNC.INVFUNC();
            inventor1.insertCore(mba.Ccs, mba.Hcs, mba.Bt, mba.Bg, mba.Hbg, mba.Beday,mba.Thongdaucapton, mba.Hr);            
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
        private void BtnDelta_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
