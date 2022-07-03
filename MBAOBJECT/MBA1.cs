using System.ComponentModel;
using System;
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
using ACADTRANSFORMER.THUVIENCAD;
using Autodesk.AutoCAD.Colors;
using System.Windows.Forms;

namespace ACADTRANSFORMER.MBAOBJECT
{
    class MBA1 : ViewModel
    {
        #region "propeties"
        private static int SoCapTon = 23;
        private Winding _winding1;
        
        private double _diameter;      //Đường kính trụ
        private double _c;             //Khoảng cách hai tâm trụ
        private double _hcs;           //Chiều cao cửa sổ mạch từ
        private double _hr = 4;        //Bề dày que thông dầu
        private double _delta = 0.27;  //Bề dày lá tôn
        private double _klr = 7650;    //khối lượng riêng       
        private double _steplab = 10;  //bước lệch đỉnh  
        private double _uv;            //điện áp vol/vòng
        private double _po;            //tổn hao không tải
        private double _poyc;          //tổn hao không tải yêu cầu
        private bool _isMultiMaterial;//
        private bool _isGiamChieuDaiRau;
        public double Po
        {
            get { return _po; }
            set
            {
                if (value != _po)
                {
                    _po = value;
                    this.OnPropertyChanged("Po");
                }
            }
        }
        
        public double Poyc
        {
            get { return _poyc; }
            set
            {
                if (value != _poyc)
                {
                    _poyc = value;
                    this.OnPropertyChanged("Poyc");
                }
            }
        }
        public double Uv
        {
            get { return _uv; }
            set
            {
                if (value != _uv)
                {
                    _uv = value;
                    this.OnPropertyChanged("Uv");
                }
            }
        }

        private double _lcuoi = 10; //chieu dài cuối lô tôn.   

        private double _ke = 0.97;        //Hệ số suy giảm bề dày tôn
        private double _ks;        //Hệ số suy giảm tiết diện
        private double _chep;      //Chiều ép tổng
        private double _s;         //Tiết diện hữu ích mặt cắt trụ
        private double _sg;        //Tiết diện hữu ích mặt cắt gông

        double[] _thongdau = new double[SoCapTon];
        private double[] _bt = new double[SoCapTon];  //Cấp trụ
        private double[] _bg = new double[SoCapTon];  //Cấp gông
        private double[] _hbg = new double[SoCapTon]; //
        private double[] _realDeltai = new double[SoCapTon]; //Be day thuc la ton cap i
        private double[] _offsetHB = new double[SoCapTon];
        private int[] _thongdaucapton = new int[SoCapTon];

        private int[] _n = new int[SoCapTon];         // Số lá từng cấp
        private double[] _beday = new double[SoCapTon];
        private double[] _duongkinh = new double[SoCapTon];//duong kinh thuc tru ton
        private int _lt = 20;                        //Hệ số làm tròn lá tôn
        #endregion
        enum cot : int
        {
            A = 0,
            B = 1,
            C = 2,
            D = 3,
            E = 4,
            F = 5,
            G = 6,
            H = 7,
            I = 8,
            J = 9,
            K = 10,
            L = 11,
            M = 12,
            N = 13,
            O = 14,
            P = 15,
            Q = 16,
            R = 17,
            S = 18,
            T = 19,
            U = 20,
            V = 21,
            W = 22,
            X = 23,
            Y = 24,
            Z = 25
        };
        #region "Field"
        public Winding Winding1
        {
            get
            {
                return _winding1;
            }
            set
            {
                if (value != _winding1)
                {
                    _winding1 = value;
                    this.OnPropertyChanged("Winding1");
                }
            }
        }
        public double[] Bt
        {
            get
            {
                this.OnPropertyChanged("Section");
                return _bt;
            }
            set
            {
                if (value != _bt)
                {
                    _bt = value;
                    this.OnPropertyChanged("Bt");
                }
            }
        }
       
        public bool IsMultiMaterial
        {
            get
            {
                return _isMultiMaterial;
            }
            set
            {
                if (value != _isMultiMaterial)
                {
                    _isMultiMaterial = value;
                    this.OnPropertyChanged("IsMultiMaterial");
                }
            }
        }
        public bool IsGiamChieuDaiRau
        {
            get
            {
                return _isGiamChieuDaiRau;
            }
            set
            {
                if (value != _isGiamChieuDaiRau)
                {
                    _isGiamChieuDaiRau = value;
                    this.OnPropertyChanged("IsGiamChieuDaiRau");
                }
            }
        }
        public double[] RealDeltai
        {
            get
            {
                return _realDeltai;
            }
            set
            {
                if (value != _realDeltai)
                {
                    _realDeltai = value;
                    this.OnPropertyChanged("RealDeltai");
                }
            }
        }
        public double[] Bg
        {
            get { return _bg; }
            set
            {
                if (value != _bg)
                {
                    _bg = value;
                    this.OnPropertyChanged("Bg");
                }
            }
        }
        public double[] Hbg
        {
            get { return _hbg; }
            set
            {
                if (value != _hbg)
                {
                    _hbg = value;
                    this.OnPropertyChanged("Hbg");
                }
            }
        }
        public int[] N
        {
            get
            {
                this.OnPropertyChanged("Beday");
                this.OnPropertyChanged("Duongkinh");
                this.OnPropertyChanged("Section");
                return _n;
            }
            set
            {
                if (value != _n)
                {
                    _n = value;
                    this.OnPropertyChanged("N");
                }
            }
        }
        public double[] Beday
        {
            get
            {
                for (int i = 0; i < _beday.Length; i++)
                {
                    _beday[i] = _realDeltai[i] * _n[i];
                }
                this.OnPropertyChanged("ChieuEp");
                return _beday;
            }
            set
            {
                if (value != _beday)
                {
                    _beday = value;
                    this.OnPropertyChanged("Beday");
                }
            }
        }
        public double[] Duongkinh
        {
            get
            {
                double bedayi = 0;
                for (int i = 0; i < _bt.Length; i++)
                {
                    if (_bt[i] != 0)
                    {
                        bedayi = bedayi + Beday[i] + _thongdau[i] * _hr;
                        _duongkinh[i] = Math.Sqrt(_bt[i] * _bt[i] + bedayi * bedayi);
                    }
                    else _duongkinh[i] = 0;

                }
                return _duongkinh;
            }
            set
            {
                if (value != _duongkinh)
                {
                    _duongkinh = value;
                    this.OnPropertyChanged("Duongkinh");
                }
            }
        }

        public double Delta
        {
            get
            {
                this.OnPropertyChanged("Section");
                this.OnPropertyChanged("Beday");
                this.OnPropertyChanged("RealDeltai");
                return _delta;
            }
            set
            {
                if (value != _delta)
                {
                    _delta = value;
                    TinhBeDayThucLaTonn();
                    this.OnPropertyChanged("Delta");
                }
            }
        }
        public double Klr
        {
            get { return _klr; }
            set
            {
                if (value != _klr)
                {
                    _klr = value;
                    this.OnPropertyChanged("Klr");
                }
            }
        }
        public double Steplab
        {
            get { return _steplab; }
            set
            {
                if (_steplab != value)
                {
                    _steplab = value;
                    this.OnPropertyChanged("Steplab");
                }
            }
        }
        public double Lcuoi
        {
            get { return _lcuoi; }
            set
            {
                if (_lcuoi != value)
                {
                    _lcuoi = value;
                    this.OnPropertyChanged("Lcuoi");
                }
            }
        }
        public double Diameter
        {
            get { return _diameter; }
            set
            {
                if (value != _diameter)
                {
                    _diameter = value;
                    this.OnPropertyChanged("Diameter");
                }
            }
        }
        public double Ccs
        {
            get { return _c; }
            set
            {
                if (value != _c)
                {
                    _c = value;
                    this.OnPropertyChanged("Ccs");
                }
            }
        }
        public double Hcs
        {
            get { return _hcs; }
            set
            {
                if (value != _hcs)
                {
                    _hcs = value;
                    this.OnPropertyChanged("Hcs");
                }
            }
        }
        public double Hr
        {
            get
            {
                this.OnPropertyChanged("ChieuEp");
                this.OnPropertyChanged("Duongkinh");
                return _hr;
            }
            set
            {
                if (value != _hr)
                {
                    _hr = value;
                    this.OnPropertyChanged("Hr");
                }
            }
        }

        public int[] Thongdaucapton
        {
            get
            {
                _thongdau[0] = _thongdaucapton[0];
                for (int i = 1; i < _thongdaucapton.Length; i++)
                {
                    _thongdau[i] = _thongdaucapton[i] * 2;
                }
                this.OnPropertyChanged("ChieuEp");
                this.OnPropertyChanged("Duongkinh");
                return _thongdaucapton;
            }
            set
            {
                if (value != _thongdaucapton)
                {
                    _thongdaucapton = value;
                    this.OnPropertyChanged("Thongdaucapton");
                }

            }
        }
        public double Ke
        {
            get
            {
                this.OnPropertyChanged("Section");
                this.OnPropertyChanged("Beday");
                this.OnPropertyChanged("RealDeltai");
                return _ke;
            }
            set
            {
                if (value != _ke)
                {
                    _ke = value;
                    TinhBeDayThucLaTonn();
                    this.OnPropertyChanged("Ke");
                }
            }
        }
        public int Lt
        {
            get { return _lt; }
            set
            {
                if (value != _lt)
                {
                    _lt = value;
                    this.OnPropertyChanged("Lt");
                }
            }
        }

        public double Ks
        {
            get
            {
                this.OnPropertyChanged("Section");
                return _ks;
            }
            set
            {
                if (value != _ks)
                {
                    _ks = value;
                    this.OnPropertyChanged("Ks");

                }
            }
        }
        public double ChieuEp
        {
            get
            {
                _chep = 0;
                for (int i = 0; i < _beday.Length; i++)
                {
                    _chep = _chep + _beday[i] + _hr * _thongdau[i];
                }
                return _chep;
            }
            set
            {
                if (value != _chep)
                {
                    _chep = value;
                    this.OnPropertyChanged("ChieuEp");
                }
            }
        }
        public double Section
        {
            get
            {
                int vol = _bt.Length;
                _s = 0;
                ;
                for (int i = 0; i < vol; i++)
                {
                    _s = _s + _n[i]*_realDeltai[i]* _bt[i]*_ks;
                }
                return _s;
            }
            set
            {
                if (value != _s)
                {
                    _s = value;
                    this.OnPropertyChanged("Section");
                }
            }
        }
        #endregion
        #region "Method"
        //Ham khoi tao
        public MBA1()
        {
            _winding1 = new Winding();
        }
        public void AutoBg()
        {

            double[] temp = new double[23] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0,0 };
            for (int i = 0; i < temp.Length; i++)
            {
                temp[i] = _bt[i];
            }
            Bg = temp;
        }
        public void AutoHbg()
        {
            int vol = _bt.Length;
            for (int i = 0; i < vol; i++)
            {
                if (_bt[i] != 0)
                    _hbg[i] = (_bg[0] - _bg[i]) / 2;
            }
            this.OnPropertyChanged("Hbg");
        }
        public void Tinhsolaton()
        {
            int[] solaton = new int[SoCapTon];
            double bedaycaptruoc = 0;            

            if (_bt[0] != 0)
            {
                solaton[0] = (int)Math.Round((Math.Sqrt(_diameter * _diameter - _bt[0] * _bt[0]) - bedaycaptruoc - _thongdau[0] * _hr) / (int)(_lt / 2) / _realDeltai[0], 0) * (int)(_lt / 2);
                bedaycaptruoc = bedaycaptruoc + solaton[0] * _realDeltai[0] + _thongdau[0] * _hr;
            }

            for (int i = 1; i < _bt.Length; i++)
            {
                if (_bt[i] != 0)
                {
                    solaton[i] = (int)Math.Round((Math.Sqrt(_diameter * _diameter - _bt[i] * _bt[i]) - bedaycaptruoc - _thongdau[i] * _hr) / _lt / _realDeltai[i], 0) * _lt;
                    bedaycaptruoc = bedaycaptruoc + solaton[i] * _realDeltai[i] + _thongdau[i] * _hr;
                }
            }
            N = solaton;
        }
        public void TinhBeDayThucLaTonn() // Tinh be day thuc cap ton i neu dung chung 1 loai ton
        {         
            if(_isMultiMaterial)
            {
                for (int i = 0; i < _bt.Length; i++)
                {
                    RealDeltai[i] = _ke * _delta;
                }
            }                   
        }
        public void vematcattru(double centerX, double centerY) //dung
        {
            Point3d center1 = new Point3d(centerX, centerY, 0);
            MyAcad.AddCircle(center1, _diameter / 2);
            int vol = _bt.Length;
            double j = centerX;
            for (int i = 0; i < vol; i++)
            {
                if (_bt[i] != 0) ve_cn2(j + _thongdau[i] / 2 * _hr, centerY, _beday[i] / 2, _bt[i]);
                if (_bt[i] != 0) ve_cn2(2 * centerX - j - _thongdau[i] / 2 * _hr, centerY, -_beday[i] / 2, _bt[i]);
                j = j + _beday[i] / 2 + _thongdau[i] / 2 * _hr;
            }
        }
        public void vematcattrungang(double centerX, double centerY) //ngang
        {
            Point3d center1 = new Point3d(centerX, centerY, 0);
            MyAcad.AddCircle(center1, _diameter / 2);

            double j = centerY;
            int vol = _bt.Length;
            for (int i = 0; i < vol; i++)
            {
                if (_bt[i] != 0) ve_cn1(centerX, j + _thongdau[i] / 2 * _hr, _bt[i], _beday[i] / 2); //nua tren
                if (_bt[i] != 0) ve_cn1(centerX, 2 * centerY - j - _thongdau[i] / 2 * _hr, _bt[i], -_beday[i] / 2); //nua duoi
                j = j + _beday[i] / 2 + _thongdau[i] / 2 * _hr;
            }
        }
        public void vematcatgong(double centerX, double centerY)
        {
            Point3d center1 = new Point3d(centerX, centerY, 0);
            MyAcad.AddCircle(center1, _diameter / 2);

            _offsetHB[0] = 0;
            int vol = _bt.Length;
            for (int i = 1; i < vol; i++)
            {
                if (_bg[i] != 0) _offsetHB[i] = (_bg[0] - _bg[i]) / 2 - _hbg[i];
            }
            double j = centerX;
            for (int i = 0; i < vol; i++)
            {
                if (_bg[i] != 0) ve_cn2(j + _thongdau[i] / 2 * _hr, centerY + _offsetHB[i], _beday[i] / 2, _bg[i]);
                if (_bg[i] != 0) ve_cn2(2 * centerX - j - _thongdau[i] / 2 * _hr, centerY + _offsetHB[i], -_beday[i] / 2, _bg[i]);
                j = j + _beday[i] / 2 + _thongdau[i] / 2 * _hr;
            }
        }
        public void vematcatgongduoi(double centerX, double centerY)
        {
            Point3d center1 = new Point3d(centerX, centerY, 0);
            MyAcad.AddCircle(center1, _diameter / 2);

            _offsetHB[0] = 0;
            int vol = _bt.Length;
            for (int i = 1; i < vol; i++)
            {
                if (_bg[i] != 0) _offsetHB[i] = (_bg[0] - _bg[i]) / 2 - _hbg[i];
            }
            double j = centerX;
            for (int i = 0; i < vol; i++)
            {
                if (_bg[i] != 0) ve_cn2(j + _thongdau[i] / 2 * _hr, centerY + _offsetHB[i], _beday[i] / 2, _bg[i]);
                if (_bg[i] != 0) ve_cn2(2 * centerX - j - _thongdau[i] / 2 * _hr, centerY + _offsetHB[i], -_beday[i] / 2, _bg[i]);
                j = j + _beday[i] / 2 + _thongdau[i] / 2 * _hr;
            }
        }
        public void vematcatgongtren(double centerX, double centerY)
        {
            Point3d center1 = new Point3d(centerX, centerY, 0);
            MyAcad.AddCircle(center1, _diameter / 2);

            _offsetHB[0] = 0;
            int vol = _bt.Length;
            for (int i = 1; i < vol; i++)
            {
                if (_bg[i] != 0) _offsetHB[i] = -(_bg[0] - _bg[i]) / 2 + _hbg[i];
            }
            double j = centerX;
            for (int i = 0; i < vol; i++)
            {
                if (_bg[i] != 0) ve_cn2(j + _thongdau[i] / 2 * _hr, centerY + _offsetHB[i], _beday[i] / 2, _bg[i]);
                if (_bg[i] != 0) ve_cn2(2 * centerX - j - _thongdau[i] / 2 * _hr, centerY + _offsetHB[i], -_beday[i] / 2, _bg[i]);
                j = j + _beday[i] / 2 + _thongdau[i] / 2 * _hr;
            }
        }
        public void vehinhchieu(double startPoint_X, double startPoint_Y) //startPoint_X,Y : tam HCD
        {
            // ---------hinh chieu dung
            //tam cua so ben trai
            double centerLeft_HCD_X = startPoint_X - _c / 2;
            double centerLeft_HCD_Y = startPoint_Y;

            //tam cua so ben phai
            double centerRight_HCD_X = startPoint_X + _c / 2;
            double centerRight_HCD_Y = startPoint_Y;
            int vol = _bt.Length;
            for (int i = 0; i < vol; i++)
            {
                if (_bt[i] != 0)
                {
                    ve_cn3(startPoint_X, startPoint_Y, 2 * _c + _bt[i], _hcs + 2 * _bg[i] + _hbg[i] * 2);
                    ve_cn3(centerLeft_HCD_X, centerLeft_HCD_Y, _c - _bt[i], _hcs + _hbg[i] * 2);
                    ve_cn3(centerRight_HCD_X, centerRight_HCD_Y, _c - _bt[i], _hcs + _hbg[i] * 2);
                }
            }

            //------- hinh chieu canh
            double center_HCC_top_X = startPoint_X + _c * 3;
            double center_HCC_top_Y = startPoint_Y + _hcs / 2 + _bg[0] / 2;
            double center_HCC_bot_X = startPoint_X + _c * 3;
            double center_HCC_bot_Y = startPoint_Y - _hcs / 2 - _bg[0] / 2;
            vematcatgongduoi(center_HCC_bot_X, center_HCC_bot_Y);  //mat cat gong tu duoi
            vematcatgongtren(center_HCC_top_X, center_HCC_top_Y);  //mat cat gong tu tren

            //ve duong noi hai mat cat gong tren va duoi            
            MyAcad.AddLine2d(center_HCC_bot_X - _chep / 2, center_HCC_bot_Y, center_HCC_top_X - _chep / 2, center_HCC_top_Y);
            MyAcad.AddLine2d(center_HCC_bot_X + _chep / 2, center_HCC_bot_Y, center_HCC_top_X + _chep / 2, center_HCC_top_Y);

            // ve hinh chieu bang
            double center_HCB_A_X = startPoint_X - _c;
            double center_HCB_A_Y = startPoint_Y - _hcs - _bt[0] * 2;
            double center_HCB_B_X = startPoint_X;
            double center_HCB_B_Y = startPoint_Y - _hcs - _bt[0] * 2;
            double center_HCB_C_X = startPoint_X + _c;
            double center_HCB_C_Y = startPoint_Y - _hcs - _bt[0] * 2;

            vematcattru(center_HCB_A_X, center_HCB_A_Y);
            vematcattru(center_HCB_B_X, center_HCB_B_Y);
            vematcattru(center_HCB_C_X, center_HCB_C_Y);
        }
        public void veboiday(double startPoint_X, double startPoint_Y) //startPoint_X,Y : tam HCD
        {
            // ---------hinh chieu dung
            //tam cua so ben trai
            double centerLeft_HCD_X = startPoint_X - _c / 2;
            double centerLeft_HCD_Y = startPoint_Y;

            //tam cua so ben phai
            double centerRight_HCD_X = startPoint_X + _c / 2;
            double centerRight_HCD_Y = startPoint_Y;
            int vol = _bt.Length;
            for (int i = 0; i < vol; i++)
            {
                if (_bt[i] != 0)
                {
                    ve_cn3(startPoint_X, startPoint_Y, 2 * _c + _bt[i], _hcs + 2 * _bg[i] + _hbg[i] * 2);
                    ve_cn3(centerLeft_HCD_X, centerLeft_HCD_Y, _c - _bt[i], _hcs + _hbg[i] * 2);
                    ve_cn3(centerRight_HCD_X, centerRight_HCD_Y, _c - _bt[i], _hcs + _hbg[i] * 2);
                }
            }

            //------- hinh chieu canh
            double center_HCC_top_X = startPoint_X + _c * 3;
            double center_HCC_top_Y = startPoint_Y + _hcs / 2 + _bg[0] / 2;
            double center_HCC_bot_X = startPoint_X + _c * 3;
            double center_HCC_bot_Y = startPoint_Y - _hcs / 2 - _bg[0] / 2;
            vematcatgongduoi(center_HCC_bot_X, center_HCC_bot_Y);  //mat cat gong tu duoi
            vematcatgongtren(center_HCC_top_X, center_HCC_top_Y);  //mat cat gong tu tren

            //ve duong noi hai mat cat gong tren va duoi            
            MyAcad.AddLine2d(center_HCC_bot_X - _chep / 2, center_HCC_bot_Y, center_HCC_top_X - _chep / 2, center_HCC_top_Y);
            MyAcad.AddLine2d(center_HCC_bot_X + _chep / 2, center_HCC_bot_Y, center_HCC_top_X + _chep / 2, center_HCC_top_Y);

            // ve hinh chieu bang
            double center_HCB_A_X = startPoint_X - _c;
            double center_HCB_A_Y = startPoint_Y - _hcs - _bt[0] * 2;
            double center_HCB_B_X = startPoint_X;
            double center_HCB_B_Y = startPoint_Y - _hcs - _bt[0] * 2;
            double center_HCB_C_X = startPoint_X + _c;
            double center_HCB_C_Y = startPoint_Y - _hcs - _bt[0] * 2;

            vematcattrungang(center_HCB_A_X, center_HCB_A_Y);
            vematcattrungang(center_HCB_B_X, center_HCB_B_Y);
            vematcattrungang(center_HCB_C_X, center_HCB_C_Y);

            Color colorW1 = Color.FromColorIndex(ColorMethod.ByAci, 4);
            Color colorW2 = Color.FromColorIndex(ColorMethod.ByAci, 3);
            Color colorW3 = Color.FromColorIndex(ColorMethod.ByAci, 1);
            Color colorW4 = Color.FromColorIndex(ColorMethod.ByAci, 6);
            Color colorW5 = Color.FromColorIndex(ColorMethod.ByAci, 5);

            //Ve boi day hinh chieu bang
            _winding1.Ve1Pha(center_HCB_A_X, center_HCB_A_Y);
            _winding1.Ve1Pha(center_HCB_B_X, center_HCB_B_Y);
            _winding1.Ve1Pha(center_HCB_C_X, center_HCB_C_Y);
            //Ve boi day HCC ben phai
            MyAcad.AddrecWidthHigh(center_HCC_bot_X + _winding1.D1t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, (_winding1.D1n - _winding1.D1t) / 2, _winding1.H1, colorW1);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X + _winding1.D2t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, (_winding1.D2n - _winding1.D2t) / 2, _winding1.H2, colorW2);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X + _winding1.D3t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, (_winding1.D3n - _winding1.D3t) / 2, _winding1.H3, colorW3);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X + _winding1.D4t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, (_winding1.D4n - _winding1.D4t) / 2, _winding1.H4, colorW4);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X + _winding1.D5t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, (_winding1.D5n - _winding1.D5t) / 2, _winding1.H5, colorW5);    //Vẽ bối dây w1
            //Ve boi day HCC ben trai
            MyAcad.AddrecWidthHigh(center_HCC_bot_X - _winding1.D1t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, -(_winding1.D1n - _winding1.D1t) / 2, _winding1.H1, colorW1);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X - _winding1.D2t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, -(_winding1.D2n - _winding1.D2t) / 2, _winding1.H2, colorW2);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X - _winding1.D3t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, -(_winding1.D3n - _winding1.D3t) / 2, _winding1.H3, colorW3);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X - _winding1.D4t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, -(_winding1.D4n - _winding1.D4t) / 2, _winding1.H4, colorW4);    //Vẽ bối dây w1
            MyAcad.AddrecWidthHigh(center_HCC_bot_X - _winding1.D5t / 2, center_HCC_bot_Y + _bg[0] / 2 + _winding1.L01, -(_winding1.D5n - _winding1.D5t) / 2, _winding1.H5, colorW5);    //Vẽ bối dây w1
        }
        public void ve_cn1(double x, double y, double with, double high) //ve hinh cn biet tam cạnh duoi và rộng, cao
        {
            MyAcad.AddrecWidthHigh(x - with / 2, y, with, high);
        }
        public void ve_cn2(double x, double y, double with, double high) //ve hinh cn biet tam cạnh trai va rong, cao
        {
            MyAcad.AddrecWidthHigh(x, y - high / 2, with, high);
        }
        public void ve_cn3(double x, double y, double with, double high) //ve hinh cn biet tam 
        {
            MyAcad.AddrecWidthHigh(x - with / 2, y - high / 2, with, high);
        }

        public void insertTableCore(double startPointX, double startPointY) //chen bang kich thuoc mach tu 
        {
            try
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
                            BlockTableRecord btr = (BlockTableRecord)acTrans.GetObject(acCurDB.CurrentSpaceId, OpenMode.ForWrite);
                            int max = _bt.Length;

                            int socaptru = 0;
                            for (int i = 0; i < max; i++)
                            {
                                if (_bt[i] != 0)
                                {
                                    socaptru = socaptru + 1;
                                }
                                else break;
                            }

                            int numRow;
                            if ((socaptru) < 15) numRow = 15; else numRow = socaptru + 3;
                            int numCol = 25;
                            double rowHeigh = 235;
                            // double textHeighData = 88;


                            //Khoi tao table
                            Table tb = new Table();
                            tb.TableStyle = acCurDB.Tablestyle;
                            tb.Position = new Point3d(startPointX, startPointY, 0);
                            tb.SetSize(numRow, numCol);
                            tb.SetRowHeight(rowHeigh);
                            tb.SetColumnWidth(400);
                            tb.Rows[0].Height = 325;
                            tb.Rows[1].Height = 620; // o chua hinh ve minh hoa
                            tb.Rows[2].Height = 360; // o Header

                            tb.Columns[1].Width = 220;// o STT
                            //Dinh dang table

                            tb.Rows[0].Style = "Data";
                            CellRange mcells0 = CellRange.Create(tb, 0, 0, 0, (int)cot.G);  //Gop 7 o hang 1
                            CellRange mcells1 = CellRange.Create(tb, 0, (int)cot.H, 0, (int)cot.L);  // Gop 7 o hang 1
                            CellRange mcells2 = CellRange.Create(tb, 1, 0, 1, (int)cot.G);  //Gop 7 o hang 2
                            CellRange mcells3 = CellRange.Create(tb, 1, (int)cot.H, 1, (int)cot.L);  //Gop 7 o hang 2
                            CellRange mcells4 = CellRange.Create(tb, 3, (int)cot.D, numRow - 1, (int)cot.D);
                            CellRange mcells5 = CellRange.Create(tb, 2, (int)cot.U, numRow - 1, (int)cot.U);
                            CellRange mcells6 = CellRange.Create(tb, 0, (int)cot.M, numRow - 1, (int)cot.M); //Gop cac o

                            tb.MergeCells(mcells0);
                            tb.MergeCells(mcells1);
                            tb.MergeCells(mcells2);
                            tb.MergeCells(mcells3);
                            tb.MergeCells(mcells4);
                            tb.MergeCells(mcells5);
                            tb.MergeCells(mcells6);

                            tb.Cells.Alignment = CellAlignment.MiddleCenter;
                            tb.Cells.TextHeight = 66;
                            tb.Columns[(int)cot.A].Width = 220;
                            tb.Columns[(int)cot.B].Width = 400;
                            tb.Columns[(int)cot.C].Width = 400;
                            tb.Columns[(int)cot.D].Width = 550;
                            tb.Columns[(int)cot.E].Width = 550;
                            tb.Columns[(int)cot.F].Width = 550;
                            tb.Columns[(int)cot.G].Width = 680;
                            tb.Columns[(int)cot.H].Width = 400;
                            tb.Columns[(int)cot.I].Width = 550;
                            tb.Columns[(int)cot.J].Width = 700;
                            tb.Columns[(int)cot.K].Width = 760;
                            tb.Columns[(int)cot.L].Width = 690;
                            tb.Columns[(int)cot.M].Width = 690;
                            tb.Columns[(int)cot.N].Width = 300;
                            tb.Columns[(int)cot.O].Width = 690;
                            tb.Columns[(int)cot.P].Width = 690;
                            tb.Columns[(int)cot.Q].Width = 690;
                            tb.Columns[(int)cot.R].Width = 690;
                            tb.Columns[(int)cot.S].Width = 800;
                            tb.Columns[(int)cot.T].Width = 400;
                            tb.Columns[(int)cot.U].Width = 400;
                            tb.Columns[(int)cot.V].Width = 800;
                            tb.Columns[(int)cot.W].Width = 1000;
                            tb.Columns[(int)cot.X].Width = 400;
                            tb.Columns[(int)cot.Y].Width = 1000;

                            tb.Rows[2].Height = 360;
                            tb.Rows[2].TextHeight = 88;
                            tb.Cells[0, (int)cot.A].TextHeight = 88;
                            tb.Cells[0, (int)cot.H].TextHeight = 88;





                            // is bit-encoded Data (1), Title (2) and Header (4)
                            //tb.SetTextHeight(textHeighData, 1);

                            // 
                            // Ghi du lieu vao table
                            //                        
                            tb.Cells[0, (int)cot.A].TextString = "XÀ NGANG VÀ TRỤ BÊN";
                            tb.Cells[0, (int)cot.H].TextString = "TRỤ GIỮA";

                            tb.Cells[2, (int)cot.A].TextString = "TT";
                            tb.Cells[2, (int)cot.B].TextString = "A";
                            tb.Cells[2, (int)cot.C].TextString = "B";
                            tb.Cells[2, (int)cot.D].TextString = "L2";
                            tb.Cells[2, (int)cot.E].TextString = "L1";
                            tb.Cells[2, (int)cot.F].TextString = "S.lá";
                            tb.Cells[2, (int)cot.G].TextString = "Chiều dày";
                            tb.Cells[2, (int)cot.H].TextString = "B";
                            tb.Cells[2, (int)cot.I].TextString = "L1";
                            tb.Cells[2, (int)cot.J].TextString = "N(5 đỉnh lệch 10)";
                            tb.Cells[2, (int)cot.K].TextString = "Chiều dày";
                            tb.Cells[2, (int)cot.L].TextString = "Chiều dài";
                            tb.Cells[2, (int)cot.N].TextString = "TT";
                            tb.Cells[2, (int)cot.O].TextString = "B.t";
                            tb.Cells[2, (int)cot.P].TextString = "B.g";
                            tb.Cells[2, (int)cot.Q].TextString = "H.bg";
                            tb.Cells[2, (int)cot.R].TextString = "N";
                            tb.Cells[2, (int)cot.S].TextString = "Bề dày";
                            tb.Cells[2, (int)cot.T].TextString = "TDx4";

                            tb.Cells[3, (int)cot.V].TextString = "Bề dày tôn";
                            tb.Cells[3, (int)cot.W].Value = _delta;
                            tb.Cells[3, (int)cot.W].DataFormat = "%lu2%pr4%zs8";

                            tb.Cells[4, (int)cot.V].TextString = "Ke";
                            tb.Cells[4, (int)cot.W].Value = _ke;
                            tb.Cells[4, (int)cot.W].DataFormat = "%lu2%pr4%zs8";

                            tb.Cells[5, (int)cot.V].TextString = "Ks";
                            tb.Cells[5, (int)cot.W].Value = _ks;
                            tb.Cells[5, (int)cot.W].DataFormat = "%lu2%pr4%zs8";

                            tb.Cells[7, (int)cot.V].TextString = "C-C";
                            tb.Cells[7, (int)cot.W].Value = Convert.ToInt16(_c);

                            tb.Cells[8, (int)cot.V].TextString = "Hcs";
                            tb.Cells[8, (int)cot.W].Value = Convert.ToInt16(_hcs);

                            tb.Cells[9, (int)cot.V].TextString = "ĐK trụ";
                            tb.Cells[9, (int)cot.W].Value = _diameter;
                            tb.Cells[9, (int)cot.W].DataFormat = "%lu2%pr1%zs8";

                            tb.Cells[10, (int)cot.V].TextString = "Chiều ép";
                            tb.Cells[10, (int)cot.V].Value = "=Sum(S4:S" + numRow.ToString() + ")+sum(T4:T" + numRow.ToString() + ")*4";

                            tb.Cells[12, (int)cot.V].TextString = "Uv";
                            tb.Cells[14, (int)cot.V].TextString = "Po(kW)";
                            //   tb.Cells[15, (int)cot.V].TextString = "Klg riêng";
                            //   tb.Cells[17, (int)cot.V].TextString = "Klg tổng";

                            //    tb.Cells[3, (int)cot.D].Contents[0].Formula = "=W8";


                            for (int i = 0; i < socaptru; i++)
                            {
                                tb.Cells[i + 3, (int)cot.A].Value = i + 1;
                                tb.Cells[i + 3, (int)cot.N].Value = i + 1;
                                tb.Cells[i + 3, (int)cot.O].Value = Convert.ToInt16(_bt[i]);
                                tb.Cells[i + 3, (int)cot.P].Value = Convert.ToInt16(_bg[i]);
                                tb.Cells[i + 3, (int)cot.Q].Value = Convert.ToInt16(_hbg[i]);
                                tb.Cells[i + 3, (int)cot.R].Value = _n[i];
                                string formularCotS = "=R" + (i + 4).ToString() + @"*$W$4*$W$5";  //Công thức tính bề dày cấp trụ
                                tb.Cells[i + 3, (int)cot.S].Contents[0].Formula = formularCotS;
                                if (i == 0) tb.Cells[3, (int)cot.T].Value = _thongdaucapton[0];
                                if (i != 0) tb.Cells[i + 3, (int)cot.T].Value = _thongdaucapton[i] * 2;
                                tb.Cells[i + 3, (int)cot.B].Contents[0].Formula = "=P" + (i + 4).ToString();
                                tb.Cells[i + 3, (int)cot.C].Contents[0].Formula = "=O" + (i + 4).ToString();
                                tb.Cells[i + 3, (int)cot.E].Contents[0].Formula = "=$W$9+2*Q" + (i + 4).ToString() + "+P" + (i + 4).ToString();
                                tb.Cells[i + 3, (int)cot.F].Contents[0].Formula = "=2*R" + (i + 4).ToString();
                                string formularCotG = "2 x %<\\AcExpr (S" + (i + 4).ToString() + "/2) \\f \"%lu2%pr2\">%";
                                tb.Cells[i + 3, (int)cot.G].Contents[0].Formula = formularCotG;
                                tb.Cells[i + 3, (int)cot.H].Contents[0].Formula = "=O" + (i + 4).ToString();
                                tb.Cells[i + 3, (int)cot.I].Contents[0].Formula = "=$W$9+2*Q" + (i + 4).ToString() + "+P" + (i + 4).ToString();
                                tb.Cells[i + 3, (int)cot.J].Contents[0].Formula = "=R" + (i + 4).ToString();
                                tb.Cells[i + 3, (int)cot.K].Contents[0].Formula = "=S" + (i + 4).ToString() + "/2";
                            }

                            //Border
                            tb.Cells.Borders.Vertical.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                            tb.Cells.Borders.Horizontal.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                            CellRange cellRange1 = CellRange.Create(tb, 0, (int)cot.M, numRow - 1, (int)cot.M);
                            tb.Cells[0, (int)cot.M].Borders.Top.IsVisible = false;
                            tb.Cells[numRow - 1, (int)cot.M].Borders.Bottom.IsVisible = false;
                            //Dinh dang du lieu trong bang
                            tb.Columns[(int)cot.K].DataFormat = "%lu2%pr2";   //Dinh dang cot K
                            tb.Columns[(int)cot.S].DataFormat = "%lu2%pr3";   //Dinh dang cot S


                            tb.GenerateLayout();
                            btr.AppendEntity(tb);
                            acTrans.AddNewlyCreatedDBObject(tb, true);
                            acTrans.Commit();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        public void insertTableCore1(double startPointX, double startPointY) //chen bang kich thuoc mach tu theo mau moi
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
                        BlockTableRecord btr = (BlockTableRecord)acTrans.GetObject(acCurDB.CurrentSpaceId, OpenMode.ForWrite);
                       // Xác định số cấp trụ tôn
                        int socaptru = 0;
                        for (int i = 0; i < _bt.Length; i++)
                        {
                            if (_bt[i] != 0)
                            {
                                socaptru = socaptru + 1;
                            }
                            else break;
                        }
                        //-----Định dạng table-------//
                        int numRow; // Số hàng của bảng mạch từ
                        if ((socaptru) < 9) numRow = 9; else numRow = socaptru +2 ;
                        int numCol = 16;
                        double rowHeigh = 210;

                        Table tb = new Table(); //Khoi tao table
                        tb.TableStyle = acCurDB.Tablestyle;
                        tb.Position = new Point3d(startPointX, startPointY, 0);
                        tb.SetSize(numRow, numCol);
                        tb.SetRowHeight(rowHeigh);
                        tb.SetColumnWidth(550);
                        
                        //Gộp các ô tiêu đề
                        tb.Rows[0].Style = "Data";

                        CellRange mcells0 = CellRange.Create(tb, 0, (int)cot.B, 0, (int)cot.D);  //Ô chứa text "TRỤ GIỮA"
                        CellRange mcells1 = CellRange.Create(tb, 0, (int)cot.E, 0, (int)cot.G);  //Ô chứa text "TRỤ BÊN"
                        CellRange mcells2 = CellRange.Create(tb, 0, (int)cot.H, 0, (int)cot.J);  //Ô chứa text "XÀ GÔNG"
                        CellRange mcells3 = CellRange.Create(tb, 0, (int)cot.K, 1, (int)cot.K);  //Ô chứa text "CHIỀU DÀY BẬC"
                        CellRange mcells4 = CellRange.Create(tb, 0, (int)cot.L, 1, (int)cot.L);  //Ô chứa text "L phôi"
                        CellRange mcells5 = CellRange.Create(tb, 0, (int)cot.M, numRow-1, (int)cot.M);  //Ô chứa text "L phôi"
                        CellRange mcells6 = CellRange.Create(tb, 0, (int)cot.A, 1, (int)cot.A);  //Ô chứa text "Bậc"
                        CellRange mcells7 = CellRange.Create(tb, 0, (int)cot.P, 1, (int)cot.P);  //Ô chứa text "Bậc"

                        tb.MergeCells(mcells0);
                        tb.MergeCells(mcells1);
                        tb.MergeCells(mcells2);
                        tb.MergeCells(mcells3);
                        tb.MergeCells(mcells4);
                        tb.MergeCells(mcells5);
                        tb.MergeCells(mcells6);
                        tb.MergeCells(mcells7);

                        tb.Cells[0,(int)cot.A].TextString = "Bậc";
                        tb.Cells[0,(int)cot.B].TextString = "Trụ giữa";
                        tb.Cells[0,(int)cot.E].TextString = "Trụ bên";
                        tb.Cells[0,(int)cot.H].TextString = "Xà gông";
                        tb.Cells[0,(int)cot.K].TextString = "Chiều dày bậc";
                        tb.Cells[0,(int)cot.L].TextString = "L phôi";
                        tb.Cells[0, (int)cot.P].TextString = "Hbg";

                        tb.Cells[1,(int)cot.B].TextString = "A";
                        tb.Cells[1,(int)cot.C].TextString = "L";
                        tb.Cells[1,(int)cot.D].TextString = "Số lá";

                        tb.Cells[1,(int)cot.E].TextString = "A";
                        tb.Cells[1,(int)cot.F].TextString = "L";
                        tb.Cells[1,(int)cot.G].TextString = "Số lá";

                        tb.Cells[1,(int)cot.H].TextString = "A";
                        tb.Cells[1,(int)cot.I].TextString = "L";
                        tb.Cells[1,(int)cot.J].TextString = "Số lá";

                        tb.Cells[0, (int)cot.N].TextString = "Bề dày tôn"; 
                        tb.Cells[1, (int)cot.N].TextString = "Ke";
                        tb.Cells[2, (int)cot.N].TextString = "Ks";
                        tb.Cells[3, (int)cot.N].TextString = "Hcs";
                        tb.Cells[4, (int)cot.N].TextString = "Ccs";
                        tb.Cells[5, (int)cot.N].TextString = "Đường kính trụ";
                        tb.Cells[6, (int)cot.N].TextString = "Uv";
                        tb.Cells[7, (int)cot.N].TextString = "Bề day rãnh dầu";
                        tb.Cells[8, (int)cot.N].TextString = "Số rãnh dầu";
                        tb.Cells[9, (int)cot.N].TextString = "Chiều ép";
                        tb.Cells[10, (int)cot.N].TextString = "Po";

                        tb.Cells[0, (int)cot.O].TextString = _delta.ToString();
                        tb.Cells[1, (int)cot.O].TextString = _ke.ToString();
                        tb.Cells[2, (int)cot.O].TextString = _ks.ToString();
                        tb.Cells[3, (int)cot.O].TextString = _hcs.ToString();
                        tb.Cells[4, (int)cot.O].TextString = _c.ToString();
                        tb.Cells[5, (int)cot.O].TextString = _diameter.ToString();
                        tb.Cells[6, (int)cot.O].TextString = _uv.ToString();
                        tb.Cells[7, (int)cot.O].TextString = _hr.ToString();

                        tb.Cells[9, (int)cot.O].TextString = _chep.ToString();
                        tb.Cells[10, (int)cot.O].TextString = _po.ToString();
                        


                        int soranhdau = 0;
                        string strRow = "";
                        for (int i = 0; i < socaptru;i++ )
                        {
                            strRow = (i + 3).ToString();
                            soranhdau = soranhdau + (int)_thongdau[i];
                            tb.Cells[i+2, (int)cot.A].TextString = (i+1).ToString();    // Đánh số thứ tự bậc
                            tb.Cells[i + 2, (int)cot.B].TextString = _bt[i].ToString(); // Điền cấp trụ giữa
                            tb.Cells[i + 2, (int)cot.C].TextString = "=$O$4+2*P" + strRow + "+H" + strRow; //Điền chiều dài lá trụ giữa
                            tb.Cells[i + 2, (int)cot.D].TextString = _n[i].ToString(); // Điền số lá trụ giữa
                            tb.Cells[i + 2, (int)cot.E].TextString = "=B" + strRow; // Điền cấp trụ bên
                            tb.Cells[i + 2, (int)cot.F].TextString = "=C" + strRow + "+H"+strRow+"-B"+strRow;// Điền chiều dài lá trụ bên
                            tb.Cells[i + 2, (int)cot.G].TextString = "=D" + strRow + "*2";// Điền số lá trụ bên
                            tb.Cells[i + 2, (int)cot.H].TextString = _bg[i].ToString(); // Điền cấp xà gông
                            tb.Cells[i + 2, (int)cot.I].TextString = "2 x " + @"%<\AcExpr ($O$5-(H"+strRow+"-B"+strRow+ ")/2) \\f \"%lu2%pr0\">%";                            
                            tb.Cells[i + 2, (int)cot.J].TextString = "=D" + strRow + "*2";// Điền số lá xà gông
                            //tb.Cells[i + 2, (int)cot.K].TextString = Math.Round(_beday[i],2).ToString();            //Điền bề dày bậc
                            tb.Cells[i + 2, (int)cot.K].TextString = "%<\\AcExpr (D"+ strRow + "*O1*O2) \\f \"%lu2%pr2\">%";            //Điền bề dày bậc
                            tb.Cells[i + 2, (int)cot.P].TextString = _hbg[i].ToString();// Điền hạ bậc gông
                            //"+(i + 3).ToString()+"

                        }
                        tb.Cells[8, (int)cot.O].TextString = soranhdau.ToString();

                        tb.Cells.Alignment = CellAlignment.MiddleCenter;
                        tb.Cells.TextHeight = 66;
                        tb.Columns[(int)cot.A].Width = 240; // độ rộng cột stt cấp tôn
                        tb.Columns[(int)cot.N].Width = 1000; // độ rộng cột stt cấp tôn
                        tb.Rows[0].Height = 240;
                        tb.Rows[1].TextHeight = 66;
                        tb.GenerateLayout();
                        btr.AppendEntity(tb);
                        acTrans.AddNewlyCreatedDBObject(tb, true);
                        acTrans.Commit();
                    }
                }
            }
        }
    }
    public class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(String propertyName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
