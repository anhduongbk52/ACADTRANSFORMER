using Autodesk.AutoCAD.Colors;

namespace ACADTRANSFORMER.MBAOBJECT
{
    class Winding:ViewModel
    {
        #region "Field"
        private int _nCanDoc;
        private double _d1t;
        private double _d2t;
        private double _d3t;
        private double _d4t;
        private double _d5t;
        private double _d1n;
        private double _d2n;
        private double _d3n;
        private double _d4n;
        private double _d5n;
        private double _h1;
        private double _h2;
        private double _h3;
        private double _h4;
        private double _h5;
        private double _demdau1;
        private double _demdau2;
        private double _demdau3;
        private double _demdau4;
        private double _demdau5;
        private double _l01;
        private double _l02;
        #endregion
        #region "Properties"
        public double D1t
        {
            get { return _d1t; }
            set
            {
                if (_d1t != value)
                {
                    _d1t = value;
                    this.OnPropertyChanged("D1t");
                }
            }
        }
        public double D2t
        {
            get { return _d2t; }
            set
            {
                if(_d2t!=value)
                {
                    _d2t = value;
                    this.OnPropertyChanged("D2t");
                }
            }
        }
        public double D3t
        {
            get { return _d3t; }
            set
            {
                if(value !=_d3t)
                {
                    _d3t = value;
                    this.OnPropertyChanged("D3t");
                }
            }
        }
        public double D4t
        {
            get { return _d4t; }
            set
            {
                if (_d4t != value)
                {
                    _d4t = value;
                    this.OnPropertyChanged("D4t");
                }
            }
        }
        public double D5t
        {
            get { return _d5t; }
            set
            {
                if(value!=_d5t)
                {
                    _d5t = value;
                    OnPropertyChanged("D5t");
                }
            }
        }
        public double D1n
        {
            get { return _d1n; }
            set
            {
                if (_d1n != value)
                {
                    _d1n = value;
                    this.OnPropertyChanged("D1n");
                }
            }
        }
        public double D2n
        {
            get { return _d2n; }
            set
            {
                if (_d2n != value)
                {
                    _d2n = value;
                    this.OnPropertyChanged("D2n");
                }
            }
        }

        public double D3n
        {
            get { return _d3n; }
            set
            {
                if (_d3n != value)
                {
                    _d3n = value;
                    this.OnPropertyChanged("D3n");
                }
            }
        }
        public double D4n
        {
            get { return _d4n; }
            set
            {
                if (_d4n != value)
                {
                    _d4n = value;
                    this.OnPropertyChanged("D4n");
                }
            }
        }
        public double D5n
        {
            get { return _d5n; }
            set
            {
                if (_d5n != value)
                {
                    _d5n = value;
                    this.OnPropertyChanged("D5n");
                }
            }
        }

        public double H1
        {
            get { return _h1; }
            set
            {
                if(value!=_h1)
                {
                    _h1 = value;
                    this.OnPropertyChanged("H1");
                }                
            }
        }
        public double H2
        {
            get { return _h2; }
            set
            {
                if (value != _h2)
                {
                    _h2 = value;
                    this.OnPropertyChanged("H2");
                }
            }
        }
        public double H3
        {
            get { return _h3; }
            set
            {
                if (value != _h3)
                {
                    _h3 = value;
                    this.OnPropertyChanged("H3");
                }
            }
        }
        public double H4
        {
            get { return _h4; }
            set
            {
                if (value != _h4)
                {
                    _h4 = value;
                    this.OnPropertyChanged("H4");
                }
            }
        }
        public double H5
        {
            get { return _h5; }
            set
            {
                if (value != _h5)
                {
                    _h5 = value;
                    this.OnPropertyChanged("H5");
                }
            }
        }
        public int NCanDoc
        {
            get { return _nCanDoc; }
            set
            {
                if(value!=_nCanDoc)
                {
                    _nCanDoc = value;
                    OnPropertyChanged("NCanDoc");
                }                
            }
        }
        public double L01
        {
            get { return _l01; }
            set
            {
                if (value != _l01)
                {
                    _l01 = value;
                    this.OnPropertyChanged("L01");
                }
            }
        }
        public double L02
        {
            get { return _l02; }
            set
            {
                if (value != _l02)
                {
                    _l02 = value;
                    this.OnPropertyChanged("L02");
                }
            }
        }
        public double Demdau1
        {
            get { return _demdau1; }
            set
            {
                if(value!=_demdau1)
                {
                    _demdau1 = value;
                    this.OnPropertyChanged("Demdau1");
                }
            }
        }
        public double Demdau2
        {
            get { return _demdau2; }
            set
            {
                if (value != _demdau2)
                {
                    _demdau2 = value;
                    OnPropertyChanged("Demdau2");
                }
            }
        }
        public double Demdau3
        {
            get { return _demdau3; }
            set
            {
                if (value != _demdau3)
                {
                    _demdau3 = value;
                    OnPropertyChanged("Demdau3");
                }
            }
        }
        public double Demdau4
        {
            get { return _demdau4; }
            set
            {
                if (value != _demdau4)
                {
                    _demdau4 = value;
                    OnPropertyChanged("Demdau4");
                }
            }
        }
        public double Demdau5
        {
            get { return _demdau5; }
            set
            {
                if (value != _demdau5)
                {
                    _demdau5 = value;
                    OnPropertyChanged("Demdau5");
                }
            }
        }

        #endregion
        #region "Method"
        public void Ve1Pha(double centerX, double centerY)   // tại tọa độ x, y
        {
            Color colorW1 = Color.FromColorIndex(ColorMethod.ByAci, 4);
            Color colorW2 = Color.FromColorIndex(ColorMethod.ByAci, 3);
            Color colorW3 = Color.FromColorIndex(ColorMethod.ByAci, 1);
            Color colorW4 = Color.FromColorIndex(ColorMethod.ByAci, 6);
            Color colorW5 = Color.FromColorIndex(ColorMethod.ByAci, 5);

            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d1t / 2, colorW1);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d1n / 2, colorW1);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d2t / 2, colorW2);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d2n / 2, colorW2);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d3t / 2, colorW3);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d3n / 2, colorW3);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d4t / 2, colorW4);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d4n / 2, colorW4);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d5t / 2, colorW5);
            THUVIENCAD.MyAcad.AddCircle(centerX, centerY, _d5n / 2, colorW5);
        }

        #endregion
    }
}
