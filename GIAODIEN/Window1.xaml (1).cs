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
                THUVIENCAD.MyAcad.InsertBlockBorderA3(toadogoc[0], toadogoc[1], 26.4);
                THUVIENCAD.MyAcad.InsertBlockTitleA3(toadogoc[0] + 215 * 26.4, toadogoc[1] + 5 * 26.4, 26.4);
                mba.insertTableCore1(toadogoc[0] + 3978 +160 , toadogoc[1] + 7444.8 - 985);
                Point2d p1 = new Point2d(toadogoc[0] + 4290, toadogoc[1] + 6710);
                THUVIENCAD.MyAcad.DrawLaGong(p1, 185, 865);
                THUVIENCAD.MyAcad.DrawTruBen(new Point2d(p1.X + 1810, p1.Y), 185, 865, true);
                THUVIENCAD.MyAcad.DrawTruGiua(new Point3d(p1.X + 4000, p1.Y + 140, 0), true);
                string mtext = "Yêu cầu kỹ thuật: ";
                THUVIENCAD.MyAcad.InserMText(toadogoc[0] + 760, toadogoc[1] + 1360, 66, mtext);
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
            //Load Cac lineType
            THUVIENCAD.StyleAutocad.LoadLineType("Continuous");
            THUVIENCAD.StyleAutocad.LoadLineType("DASHED");
            THUVIENCAD.StyleAutocad.LoadLineType("CENTER2");
            THUVIENCAD.StyleAutocad.LoadLineType("DASHDOT");
            //Khởi tạo layer
            THUVIENCAD.StyleAutocad.CreatLayer("Visible (ISO)","Continuous", 7, LineWeight.LineWeight030, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Visible Narrow (ISO)", "Continuous", 4, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Title (ISO)", "Continuous", 7, LineWeight.LineWeight025, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Symbol (ISO)", "Continuous", 7, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("NoPlot", "Continuous", 8, LineWeight.ByLineWeightDefault, false);
            THUVIENCAD.StyleAutocad.CreatLayer("Hidden (ISO)", "DASHED", 4, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Hatch (ISO)", "Continuous", 8, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Dimension (ISO)", "Continuous", 3, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Center Mark (ISO)", "CENTER2", 2, LineWeight.LineWeight018, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Centerline (ISO)", "DASHDOT", 2, LineWeight.LineWeight015, true);
            THUVIENCAD.StyleAutocad.CreatLayer("Border (ISO)", "Continuous", 7, LineWeight.LineWeight035, true);
            //Khởi tạo text style
            THUVIENCAD.MyAcad.CreatTextStyle("STANDARD_DUONG", "TCVN 7284.ttf",0);
            THUVIENCAD.MyAcad.CreatTextStyle("X1_DUONG", "TCVN 7284.ttf", 2.5);
            THUVIENCAD.MyAcad.CreatTextStyle("X2_DUONG", "TCVN 7284.ttf", 5);
            THUVIENCAD.MyAcad.CreatTextStyle("X4_DUONG", "TCVN 7284.ttf", 10);
            THUVIENCAD.MyAcad.CreatTextStyle("X5_DUONG", "TCVN 7284.ttf", 12.5);
            THUVIENCAD.MyAcad.CreatTextStyle("X10_DUONG", "TCVN 7284.ttf", 25);
            THUVIENCAD.MyAcad.CreatTextStyle("X15_DUONG", "TCVN 7284.ttf", 37.5);
            THUVIENCAD.MyAcad.CreatTextStyle("X20_DUONG", "TCVN 7284.ttf", 50);
            THUVIENCAD.MyAcad.CreatTextStyle("X26.4_DUONG","TCVN 7284.ttf", 66);
            THUVIENCAD.MyAcad.CreatTextStyle("X36_DUONG", "TCVN 7284.ttf", 90);
            THUVIENCAD.MyAcad.CreatTextStyle("X40_DUONG", "TCVN 7284.ttf", 100);
            
            THUVIENCAD.MyAcad.CreatCostumProperties("Vẽ", "Đỗ Ánh Dương");
            THUVIENCAD.MyAcad.CreatCostumProperties("Thiết kế", "Đỗ Ánh Dương");
            THUVIENCAD.MyAcad.CreatCostumProperties("Kiểm soát", "Lê Hải Quân");
            THUVIENCAD.MyAcad.CreatCostumProperties("Ban thiết kế", "Ng.Quang Tuệ");
            THUVIENCAD.MyAcad.CreatCostumProperties("Duyệt", "Ng.Vũ Cường");
            THUVIENCAD.MyAcad.CreatCostumProperties("DateTime", "01-03-2016");
            THUVIENCAD.MyAcad.CreatCostumProperties("Công suất", "63M");
            THUVIENCAD.MyAcad.CreatCostumProperties("Điện áp", "115/38,5/23");
            THUVIENCAD.MyAcad.CreatCostumProperties("Trạm", "TIỀN TRUNG + PHÚC ĐIỀN");
            THUVIENCAD.MyAcad.CreatCostumProperties("MS", "164735-233+234");
            //Khởi tạo các dimension style
            THUVIENCAD.MyAcad.CreatDimStyle("X1_DUONG", "X1_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X2_DUONG", "X2_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X4_DUONG", "X4_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X5_DUONG", "X5_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X10_DUONG", "X10_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X15_DUONG", "X15_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X20_DUONG", "X20_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X26.4_DUONG", "X26.4_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X36_DUONG", "X36_DUONG");
            THUVIENCAD.MyAcad.CreatDimStyle("X40_DUONG", "X40_DUONG"); 
        }
      
        private void reDraw_Click(object sender, RoutedEventArgs e)
        {
           // string[] arrColor = new String[] { "#ff4e50", "#91fb66", "#ff4e50", "#ffc125", "#363636", "#a2bae0", "#000000", "#88b200", "ffc125", "b2b2b2", "#ff4e50", "#a2bae0", "#91fb66", "#ff4e50", "#000000", "#a2bae0", "#000000", "#88b200", "ffc125", "#ffc125", "#363636", "#a2bae0", "#000000", "#88b200", "ffc125", "#ff4e50", "#91fb66", "#ff4e50" };
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
                 X2 = khungve.Width,   Y2 = khungve.Width/2
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

                 System.Windows.Media.Color mycolor = Color.FromRgb(0,0,0);
                 if (mba.Bg[i] != 0)
                 {
                     GrapWPF.drawRecCenterLeft(j + ((double)mba.Thongdaucapton[i] * mba.Hr) * factor, y0 + _offsetHB[i] * factor, (mba.Beday[i] / 2) * factor, mba.Bg[i] * factor, 0.5, mycolor, khungve);
                     GrapWPF.drawRecCenterRight(2*x0-(j + ((double)mba.Thongdaucapton[i] * mba.Hr) * factor), y0 + _offsetHB[i] * factor, (mba.Beday[i] / 2) * factor, mba.Bg[i] * factor, 0.5, mycolor, khungve);
                     j = j + (mba.Beday[i] / 2 + (double)mba.Thongdaucapton[i] * mba.Hr) * factor;
                 }              
             }            
        }

    }
}
