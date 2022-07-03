using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACADTRANSFORMER.MBAOBJECT
{
    class SODOTRAI
    {
        #region "field"
        private int _soCanDoc;
        private int _soGalet;
        private double _khoangCachHaiCanDoc;
        private double _khoangCachHaiGalet;
        private double _chieudaicandoc;
        private List<Line> _listCanDoc;
        private List<Line> _listgalet;
        Point3d _centerPoint;
        #endregion
        #region "Properties"
        public int SoCanDoc { get => _soCanDoc; set => _soCanDoc = value; }
        public int SoGalet { get => _soGalet; set => _soGalet = value; }
        public double KhoangCachHaiCanDoc { get => _khoangCachHaiCanDoc; set => _khoangCachHaiCanDoc = value; }
        public double KhoangCachHaiGalet { get => _khoangCachHaiGalet; set => _khoangCachHaiGalet = value; }
        public double Chieudaicandoc { get => _chieudaicandoc; set => _chieudaicandoc = value; }
        public Line CenterLine
        {
            get
            {
                return new Line(
                    new Point3d(CenterPoint.X, CenterPoint.Y - _chieudaicandoc / 2 - _khoangCachHaiCanDoc, 0),
                    new Point3d(CenterPoint.X, CenterPoint.Y + _chieudaicandoc / 2 + _khoangCachHaiCanDoc, 0)
                    );
            }
        }
        public List<Line> ListCanDoc { get => _listCanDoc; set => _listCanDoc = value; }
        public Point3d CenterPoint { get => _centerPoint; set => _centerPoint = value; }
        public List<Line> Listgalet { get => _listgalet; set => _listgalet = value; }
        #endregion



        //Cac ham cuc bo
        private List<Line> CreatListCanDoc(double khoangCachHaiCanDoc,double chieuDaiCanDoc, int soCanDoc)
        {
            Point3d startPoint_i;
            Point3d endPoint_i;
            List<Line> ketqua = new List<Line>();
            for (int i=0;i< soCanDoc + 1;i++)
            {
                startPoint_i = new Point3d(CenterPoint.X -(khoangCachHaiCanDoc/2)- (_soCanDoc-2)/2*khoangCachHaiCanDoc +i*khoangCachHaiCanDoc, CenterPoint.Y- chieuDaiCanDoc / 2,0);
                endPoint_i =   new Point3d(CenterPoint.X -(khoangCachHaiCanDoc/2)- (_soCanDoc-2)/2*khoangCachHaiCanDoc +i*khoangCachHaiCanDoc, CenterPoint.Y + chieuDaiCanDoc / 2,0);                
                ketqua.Add(new Line(startPoint_i, endPoint_i));
            }
            return ketqua;
        }
        private List<Line> CreatListGalet(double khoangCachHaiGalet, int soGalet)
        {
            Point3d startPoint_i;
            Point3d endPoint_i;            
            List<Line> ketqua = new List<Line>();
            for (int i = 0; i < soGalet; i++)
            {
                startPoint_i = new Point3d(CenterPoint.X - (_khoangCachHaiCanDoc / 2) - (_soCanDoc - 2) / 2 * _khoangCachHaiCanDoc , CenterPoint.Y - khoangCachHaiGalet / 2-(_soGalet/2-1)*khoangCachHaiGalet+i* khoangCachHaiGalet, 0);
                endPoint_i =   new Point3d(CenterPoint.X + (_khoangCachHaiCanDoc / 2) + ((_soCanDoc - 2) / 2+1)* _khoangCachHaiCanDoc , CenterPoint.Y - khoangCachHaiGalet / 2-(_soGalet/2-1)*khoangCachHaiGalet+i* khoangCachHaiGalet, 0);
                ketqua.Add(new Line(startPoint_i, endPoint_i));
            }
            return ketqua;
        }

        public void VeCanDoc(string layerName, Extents2d zone)
        {
            //Cau hinh
            _khoangCachHaiCanDoc = Math.Abs(zone.MinPoint.X - zone.MaxPoint.X) / (_soCanDoc + 2);           
            _chieudaicandoc = Math.Abs(zone.MinPoint.Y - zone.MaxPoint.Y);
            //Ve duong tam

            //Khoi tao can doc
            _listCanDoc =  CreatListCanDoc(_khoangCachHaiCanDoc, _chieudaicandoc,_soCanDoc);            
           
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if(acDoc!=null&& _listCanDoc.Count>2)
            {
                ACADLIBRARY.AcadBase.AddLine(CenterLine, acDoc, layerName);
                foreach (var candoc in _listCanDoc)
                {
                    ACADLIBRARY.AcadBase.AddLine(candoc, acDoc, layerName);

                }
            }           
        }
        public void VeGalet(string layerName, Extents2d zone)
        {
            //Cau hinh
            _khoangCachHaiGalet = Math.Abs(zone.MinPoint.Y - zone.MaxPoint.Y) / (_soGalet + 20);
            //Ve duong tam

            _listgalet = CreatListGalet(_khoangCachHaiGalet, _soGalet);

            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (acDoc != null && _listgalet.Count > 2)
            {
                ACADLIBRARY.AcadBase.AddLine(CenterLine, acDoc, layerName);
                foreach (var galet in _listgalet)
                {
                    ACADLIBRARY.AcadBase.AddLine(galet, acDoc, layerName);

                }
            }
        }
    }
}
