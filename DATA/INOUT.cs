using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ACADTRANSFORMER.MBAOBJECT;
using System.Xml;
using Microsoft.Office.Interop;
namespace ACADTRANSFORMER.DATA
{
    class INOUT
    {
        public static bool CreateExcelAppIfNotOpen = false;
        public static void ExportDataToExcel(MBA1 m1)
        {
            excel.Application myExcelApp = null;
            excel.Workbook myExcelWorkbook = null;
            excel.Worksheet myExcelSheet = null;
            excel.Worksheet myExcelSheet1=null;
            excel.Worksheet myExcelSheet2 = null;
            //Mở App, nếu chưa thì tạo mới App và thêm workbooks.
            try
            {
                myExcelApp = (excel.Application)Marshal.GetActiveObject("Excel.Application");
                myExcelApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Lỗi: Excel chưa được mở");
                CreateExcelAppIfNotOpen = true;
            }
            if (CreateExcelAppIfNotOpen)
            {
                try
                {
                    myExcelApp = new excel.Application();
                    myExcelApp.Visible = true;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            try
            {
                if (myExcelApp != null)
                {
                    myExcelWorkbook = myExcelApp.Workbooks.Add();
                    myExcelSheet = myExcelWorkbook.Sheets.Add();
                    myExcelSheet1 = myExcelWorkbook.Sheets.Add();
                    myExcelSheet2 = myExcelWorkbook.Sheets.Add();
                    myExcelSheet.Name = "Truton";
                    myExcelSheet1.Name = "Pha phôi";
                    myExcelSheet2.Name = "BangKe";
                    //Các kích thước cơ bản
                    myExcelSheet.get_Range("A1").Value2 = "Lcs"; myExcelSheet.get_Range("A2").Value2 = m1.Hcs;
                    myExcelSheet.get_Range("B1").Value2 = "Ccs"; myExcelSheet.get_Range("B2").Value2 = m1.Ccs;
                    myExcelSheet.get_Range("C1").Value2 = "φ"; myExcelSheet.get_Range("C2").Value2 = m1.Diameter;
                    myExcelSheet.get_Range("D1").Value2 = "δ"; myExcelSheet.get_Range("D2").Value2 = m1.Delta;
                    myExcelSheet.get_Range("E1").Value2 = "ke"; myExcelSheet.get_Range("E2").Value2 = m1.Ke;
                    myExcelSheet.get_Range("F1").Value2 = "chiahet"; myExcelSheet.get_Range("F2").Value2 = m1.Lt;
                    myExcelSheet.get_Range("G1").Value2 = "hr"; myExcelSheet.get_Range("G2").Value2 = m1.Hr;
                    myExcelSheet.get_Range("H1").Value2 = "Ks"; myExcelSheet.get_Range("H2").Value2 = m1.Ks;
                    // Thông số bổ sung:
                    myExcelSheet.get_Range("N1").Value2 = "Uv"; myExcelSheet.get_Range("O1").Value2 = m1.Uv;
                    myExcelSheet.get_Range("N2").Value2 = "Po"; myExcelSheet.get_Range("O2").Value2 = m1.Po;
                    myExcelSheet.get_Range("N3").Value2 = "Poyc"; myExcelSheet.get_Range("O3").Value2 = m1.Poyc;

                    //Header 
                    myExcelSheet.get_Range("A4").Value2 = "STT";
                    myExcelSheet.get_Range("B4").Value2 = "Bt";
                    myExcelSheet.get_Range("C4").Value2 = "Bg";
                    myExcelSheet.get_Range("D4").Value2 = "Hbg";
                    myExcelSheet.get_Range("E4").Value2 = "N";
                    myExcelSheet.get_Range("F4").Value2 = "Beday";
                    myExcelSheet.get_Range("G4").Value2 = "TD";

                    myExcelSheet.get_Range("J1").Value2 = "Dt";
                    myExcelSheet.get_Range("K1").Value2 = "Dn";
                    myExcelSheet.get_Range("L1").Value2 = "H";

                    myExcelSheet.get_Range("I2").Value2 = "W1"; myExcelSheet.get_Range("J2").Value2 = m1.Winding1.D1t;
                    myExcelSheet.get_Range("I3").Value2 = "W2"; myExcelSheet.get_Range("J3").Value2 = m1.Winding1.D2t;
                    myExcelSheet.get_Range("I4").Value2 = "W3"; myExcelSheet.get_Range("J4").Value2 = m1.Winding1.D3t;
                    myExcelSheet.get_Range("I5").Value2 = "W4"; myExcelSheet.get_Range("J5").Value2 = m1.Winding1.D4t;
                    myExcelSheet.get_Range("I6").Value2 = "W5"; myExcelSheet.get_Range("J6").Value2 = m1.Winding1.D5t;
                    myExcelSheet.get_Range("I7").Value2 = "Số căn dọc"; myExcelSheet.get_Range("J7").Value2 = m1.Winding1.NCanDoc;
                    myExcelSheet.get_Range("I8").Value2 = "L01"; myExcelSheet.get_Range("J8").Value2 = m1.Winding1.L01;
                    myExcelSheet.get_Range("I9").Value2 = "L02"; myExcelSheet.get_Range("J9").Value2 = m1.Winding1.L02;

                    //Đường kính bối dây
                    myExcelSheet.get_Range("K2").Value2 = m1.Winding1.D1n;
                    myExcelSheet.get_Range("K3").Value2 = m1.Winding1.D2n;
                    myExcelSheet.get_Range("K4").Value2 = m1.Winding1.D3n;
                    myExcelSheet.get_Range("K5").Value2 = m1.Winding1.D4n;
                    myExcelSheet.get_Range("K6").Value2 = m1.Winding1.D5n;

                    myExcelSheet.get_Range("L2").Value2 = m1.Winding1.H1;
                    myExcelSheet.get_Range("L3").Value2 = m1.Winding1.H2;
                    myExcelSheet.get_Range("L4").Value2 = m1.Winding1.H3;
                    myExcelSheet.get_Range("L5").Value2 = m1.Winding1.H4;
                    myExcelSheet.get_Range("L6").Value2 = m1.Winding1.H5;

                    //
                    int numRow = m1.Bt.Length;
                    int numCol = m1.Bg.Length;
                    for (int i = 0; i < numRow; i++)
                    {
                        myExcelSheet.get_Range("A" + (i + 5).ToString()).Value2 = i + 1;  // danh so thu tu
                        myExcelSheet.get_Range("B" + (i + 5).ToString()).Value2 = m1.Bt[i];  //Ghi bề rộng cấp tôn
                        myExcelSheet.get_Range("C" + (i + 5).ToString()).Value2 = m1.Bg[i];    //Ghi bề rộng gông
                        myExcelSheet.get_Range("D" + (i + 5).ToString()).Value2 = m1.Hbg[i];      // Ghi hạ bậc
                        myExcelSheet.get_Range("E" + (i + 5).ToString()).Value2 = m1.N[i];    //Ghi số lá tôn
                        myExcelSheet.get_Range("F" + (i + 5).ToString()).Value2 = m1.Beday[i];   //Ghi bề dày cấp tôn
                        myExcelSheet.get_Range("G" + (i + 5).ToString()).Value2 = m1.Thongdaucapton[i];//ghi thông dầu
                    }

                    //ghi vào sheet pha phôi
                    myExcelSheet1.get_Range("A1").Value2 = "=TODAY()";
                    myExcelSheet1.get_Range("A2").Value2 = "Hcs"; myExcelSheet1.get_Range("A3").Value2 = m1.Hcs;
                    myExcelSheet1.get_Range("B2").Value2 = "Ccs"; myExcelSheet1.get_Range("B3").Value2 = m1.Ccs;
                    myExcelSheet1.get_Range("C2").Value2 = "D"; myExcelSheet1.get_Range("C3").Value2 = m1.Delta;
                    myExcelSheet1.get_Range("D2").Value2 = "Ke"; myExcelSheet1.get_Range("D3").Value2 = m1.Ke;
                    myExcelSheet1.get_Range("E2").Value2 = "h"; myExcelSheet1.get_Range("E3").Value2 = m1.Steplab;
                    myExcelSheet1.get_Range("F2").Value2 = "Klr"; myExcelSheet1.get_Range("F3").Value2 = m1.Klr;
                    myExcelSheet1.get_Range("G2").Value2 = "L cuối"; myExcelSheet1.get_Range("G3").Value2 = m1.Lcuoi;
                    myExcelSheet1.get_Range("H2").Value2 = "Ks"; myExcelSheet1.get_Range("H3").Value2 = m1.Ks;

                    myExcelSheet1.get_Range("A4").Value2 = "TT";
                    myExcelSheet1.get_Range("B4").Value2 = "Bt";
                    myExcelSheet1.get_Range("C4").Value2 = "Bg";
                    myExcelSheet1.get_Range("D4").Value2 = "Hbg";
                    myExcelSheet1.get_Range("E4").Value2 = "N trụ giữa";
                    myExcelSheet1.get_Range("F4").Value2 = "Chiều dày";
                    myExcelSheet1.get_Range("G4").Value2 = "L trụ";
                    myExcelSheet1.get_Range("H4").Value2 = "L gông";
                    myExcelSheet1.get_Range("I4").Value2 = "L phôi trụ";
                    myExcelSheet1.get_Range("J4").Value2 = "L phôi gông";
                    myExcelSheet1.get_Range("K4").Value2 = "L phôi";
                    myExcelSheet1.get_Range("L4").Value2 = "G phôi";
                    myExcelSheet1.get_Range("M4").Value2 = "G trụ-tinh";
                    myExcelSheet1.get_Range("N4").Value2 = "G gông-tinh";
                    myExcelSheet1.get_Range("O4").Value2 = "G tinh";
                    int vol = 0;
                    for (int i = 0; i < m1.Bt.Length;i++ )
                    {
                        if (m1.Bt[i] != 0)
                            vol = vol + 1;
                    }
                    for (int i=0;i<vol;i++)
                    {
                        string rowIndex = (i + 5).ToString();
                        myExcelSheet1.get_Range("A" + rowIndex).Value2 = i + 1;  // danh so thu tu
                        myExcelSheet1.get_Range("B" + rowIndex).Value2 = m1.Bt[i];  //Ghi bề rộng cấp tôn trụ
                        myExcelSheet1.get_Range("C" + rowIndex).Value2 = m1.Bg[i];    //Ghi bề rộng gông
                        myExcelSheet1.get_Range("D" + rowIndex).Value2 = m1.Hbg[i];      // Ghi hạ bậc
                        myExcelSheet1.get_Range("E" + rowIndex).Value2 = m1.N[i];    //Ghi số lá tôn
                        myExcelSheet1.get_Range("F" + rowIndex).Value2 = m1.Beday[i];   //Ghi bề dày cấp tôn
                        myExcelSheet1.get_Range("G" + rowIndex).Value2 = "=$A$3+D" + rowIndex + "*2+C" + rowIndex;   //Ghi chiều dài lá trụ
                        myExcelSheet1.get_Range("H" + rowIndex).Value2 = "=$B$3*2";   //Ghi chiều dài lá gông
                        myExcelSheet1.get_Range("I" + rowIndex).Value2 = "=(3*G" + rowIndex + "*E" + rowIndex + "-2*$E$3*(6/5*E" + rowIndex+"-3))/1000";   //Ghi chiều dài phôi trụ
                        myExcelSheet1.get_Range("J" + rowIndex).Value2 = "=(2*H" + rowIndex + "*E" + rowIndex + ")/1000";   //Ghi chiều dài phôi gông
                        myExcelSheet1.get_Range("K" + rowIndex).Value2 = "=$G$3+I" + rowIndex + "+SUMIF($C$5:$C$" + (vol + 4).ToString() + ",B" + rowIndex + ",$J$5:$J$" + (vol + 4).ToString() + ")";   //Ghi chiều dài phôi tổng
                        myExcelSheet1.get_Range("L" + rowIndex).Value2 = "=$C$3*$D$3*$H$3*$F$3*B" + rowIndex + "*K" + rowIndex + "/1000000";   //Ghi khối lượng phôi tổng                  
                        myExcelSheet1.get_Range("M" + rowIndex).Value2 = "=$F$3*$D$3*$H$3*$C$3/1000000000*(3*G" + rowIndex + "*B" + rowIndex + "-B" + rowIndex + "^2/2-4*$E$3^2)*E" + rowIndex;   //Ghi khối lượng trụ tinh                  
                        myExcelSheet1.get_Range("N" + rowIndex).Value2 = "=$F$3*$D$3*$H$3*$C$3/1000000000*(C" + rowIndex + "*2*(H" + rowIndex + "-C" + rowIndex + "/4)*E" + rowIndex + ")";   //Ghi khối lượng gông tinh                  
                        myExcelSheet1.get_Range("O" + rowIndex).Value2 = "=M" + rowIndex + "+N" + rowIndex;   //G tinh
                    }
                    myExcelSheet1.get_Range("L" + (vol + 5).ToString()).Value2 = "=SUM(L5:L" + (vol + 4).ToString() + ")"; //Ghi khối lượng phôi tổng
                    myExcelSheet1.get_Range("O" + (vol + 5).ToString()).Value2 = "=SUM(O5:O" + (vol + 4).ToString() + ")"; //Ghi khối lượng tinh tổng

                    //Ghi vao sheet bảng kê tôn
                    myExcelSheet2.PageSetup.LeftHeader = "MBA 63MVA 115/38,5/23kV";
                    myExcelSheet2.PageSetup.RightHeader = "Trạm Phúc Điền";
                    
                    myExcelSheet2.get_Range("A4").EntireColumn.ColumnWidth= 5;
                    myExcelSheet2.get_Range("B4").EntireColumn.ColumnWidth = 12;
                    myExcelSheet2.get_Range("C4").EntireColumn.ColumnWidth = 14;
                    myExcelSheet2.get_Range("D4").EntireColumn.ColumnWidth = 18;
                    myExcelSheet2.get_Range("E4").EntireColumn.ColumnWidth = 18;
                    myExcelSheet2.get_Range("F4").EntireColumn.ColumnWidth = 18;

                    myExcelSheet2.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    myExcelSheet2.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    myExcelSheet2.Cells.Font.Size = 12;
                    myExcelSheet2.Cells.Font.Name = "Times New Roman";
                    myExcelSheet2.get_Range("A1", "F4").Cells.Font.Bold = true;
                    myExcelSheet2.get_Range("A1", "F3 ").Cells.Font.Size = 16;
                    myExcelSheet2.get_Range("A4", "F5").Cells.WrapText = true;
                    myExcelSheet2.get_Range("A1", "F1").Merge(true);
                    myExcelSheet2.get_Range("A2", "F2").Merge(true);
                    myExcelSheet2.get_Range("A3", "F3").Merge(true);
                    myExcelSheet2.get_Range("C5", "C"+(vol+4).ToString()).Merge(true);

                    myExcelSheet2.get_Range("A1").Value2 = "BẢNG KÊ KHỐI LƯỢNG TÔN";
                    myExcelSheet2.get_Range("A2").Value2 = "MÁY BIẾN ÁP 63MVA 115/38,5/23 KV";
                    myExcelSheet2.get_Range("A3").Value2 = "TRẠM PHÚC ĐIỀN";
                    myExcelSheet2.get_Range("A4").Value2 = "TT";
                    myExcelSheet2.get_Range("B4").Value2 = "Khổ tôn";
                    myExcelSheet2.get_Range("C4").Value2 = "Quy cách";
                    myExcelSheet2.get_Range("D4").Value2 = "Chiều dài /1 máy /n (m)";
                    myExcelSheet2.get_Range("E4").Value2 = "Khối lượng /1 máy /n (kG)";
                    myExcelSheet2.get_Range("F4").Value2 = "Ghi chú";

                    for (int i = 0; i < vol; i++)
                    {
                        string rowIndex = (i + 5).ToString();
                        myExcelSheet2.get_Range("A" + rowIndex).Value2 = i + 1;  // danh so thu tu
                        myExcelSheet2.get_Range("B" + rowIndex).Value2 = m1.Bt[i];  //Ghi bề rộng cấp tôn trụ
                        myExcelSheet2.get_Range("D" + rowIndex).Value2 = "=ROUND('Pha phôi'!K"+rowIndex+"*$G$3/10,0)*10";  //Ghi chieu dai dat mua
                        myExcelSheet2.get_Range("E" + rowIndex).Value2 = "=ROUND('Pha phôi'!L"+rowIndex+"*$G$3/10,0)*10";  //Ghi bề khoi luong dat mua
                    }
                    myExcelSheet2.get_Range("G2").Value2 = "Dự phòng";
                    myExcelSheet2.get_Range("G3").Value2 = 1.04;
                    myExcelSheet2.get_Range("E" + (vol + 5).ToString()).Value2 = "=SUM(E5:E"+(vol+4).ToString()+")";
                    myExcelSheet2.get_Range("A" + (vol + 5).ToString()).Value2 = "Tổng";
                    myExcelSheet2.get_Range("A" + (vol + 5).ToString()).Font.Bold = true;
                    myExcelSheet2.get_Range("E" + (vol + 5).ToString()).Font.Bold = true;
                    myExcelSheet2.get_Range("C5:C" + (vol + 4).ToString()).Merge(false);
                    myExcelSheet2.get_Range("C5").Value2 = "Tôn nhật dày 0,23 suất tổn hao 0,9w/kg ";
                    myExcelSheet2.get_Range("B" + (vol + 7).ToString() + ":C" + (vol + 7).ToString()).Merge(true);
                    myExcelSheet2.get_Range("E" + (vol + 6).ToString() + ":F" + (vol + 6).ToString()).Merge(true);
                    myExcelSheet2.get_Range("E" + (vol + 7).ToString() + ":F" + (vol + 7).ToString()).Merge(true);
                    myExcelSheet2.get_Range("A" + (vol + 5).ToString() + ":D" + (vol + 5).ToString()).Merge(true);
                    
                    myExcelSheet2.get_Range("E" + (vol + 6).ToString()).Value2 = "Ngày      tháng       năm     ";
                    myExcelSheet2.get_Range("B" + (vol + 7).ToString()).Value2 = "Duyệt";
                    myExcelSheet2.get_Range("D" + (vol + 7).ToString()).Value2 = "Ban thiết kế";
                    myExcelSheet2.get_Range("E" + (vol + 7).ToString()).Value2 = "Người lập";
                    myExcelSheet2.get_Range("A" + (vol + 7).ToString() + ":F" + (vol + 7).ToString()).Font.Bold=true;
                    myExcelSheet2.get_Range("A4:F" + (vol + 5).ToString()).Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    myExcelSheet2.get_Range("A4:F" + (vol + 5).ToString()).Borders.Weight = 2d;
                    myExcelSheet2.PageSetup.PrintArea = "$A$1:$F$24";
                    //Microsoft.Office.Interop.Excel.Range headerRange = myExcelSheet2.get_Range("A1", "F1");
                    //headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //headerRange.Value = "MBA 63MVA 115/38,5/23kV";
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static MBA1 ImportDataFromExcel(string path)
        {
            MBA1 m1 = new MBA1();
            excel.Application myExcelApp = new excel.Application();
            excel.Workbook myExcelWorkbook = myExcelApp.Workbooks.Open(path, ReadOnly: true);
            try
            {      
                excel.Worksheet myExcelSheet = myExcelWorkbook.Worksheets["Truton"];
                // Read
                m1.Hcs = myExcelSheet.get_Range("A2").Value2;
                m1.Ccs = myExcelSheet.get_Range("B2").Value2;
                m1.Diameter = myExcelSheet.get_Range("C2").Value2;
                m1.Delta = myExcelSheet.get_Range("D2").Value2;
                m1.Ke = myExcelSheet.get_Range("E2").Value2;
                m1.Lt = (int)myExcelSheet.get_Range("F2").Value2;
                m1.Hr = myExcelSheet.get_Range("G2").Value2;
                m1.Ks = myExcelSheet.get_Range("H2").Value2;

                m1.Uv = myExcelSheet.get_Range("O1").Value2;
                m1.Po = myExcelSheet.get_Range("O2").Value2;
                m1.Poyc = myExcelSheet.get_Range("O3").Value2;

                m1.Winding1.D1t = myExcelSheet.get_Range("J2").Value2;
                m1.Winding1.D2t = myExcelSheet.get_Range("J3").Value2;
                m1.Winding1.D3t = myExcelSheet.get_Range("J4").Value2;
                m1.Winding1.D4t = myExcelSheet.get_Range("J5").Value2;
                m1.Winding1.D5t = myExcelSheet.get_Range("J6").Value2;

                m1.Winding1.D1n = myExcelSheet.get_Range("K2").Value2;
                m1.Winding1.D2n = myExcelSheet.get_Range("K3").Value2;
                m1.Winding1.D3n = myExcelSheet.get_Range("K4").Value2;
                m1.Winding1.D4n = myExcelSheet.get_Range("K5").Value2;
                m1.Winding1.D5n = myExcelSheet.get_Range("K6").Value2;

                m1.Winding1.H1 = myExcelSheet.get_Range("L2").Value2;
                m1.Winding1.H2 = myExcelSheet.get_Range("L3").Value2;
                m1.Winding1.H3 = myExcelSheet.get_Range("L4").Value2;
                m1.Winding1.H4 = myExcelSheet.get_Range("L5").Value2;
                m1.Winding1.H5 = myExcelSheet.get_Range("L6").Value2;

                for (int i = 0; i < m1.Bt.Length; i++)
                {
                    m1.Bt[i] = myExcelSheet.get_Range("B" + (i + 5).ToString()).Value2;
                    m1.Bg[i] = myExcelSheet.get_Range("C" + (i + 5).ToString()).Value2;
                    m1.Hbg[i] = myExcelSheet.get_Range("D" + (i + 5).ToString()).Value2;
                    m1.N[i] = (int)myExcelSheet.get_Range("E" + (i + 5).ToString()).Value2;
                    m1.Beday[i] = myExcelSheet.get_Range("F" + (i + 5).ToString()).Value2;
                    m1.Thongdaucapton[i] = (int)myExcelSheet.get_Range("G" + (i + 5).ToString()).Value2;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                object misValue = System.Reflection.Missing.Value;
                myExcelWorkbook.Close(false, misValue, misValue);

                myExcelApp.Quit();
            }            
            return m1;
        }
        public static void SaveToFileXml(MBA1 m1,string path) // tao mot file xml
        {

            XmlDocument xmlDoc = new XmlDocument();
            // Create a root element
            XmlElement rootname = xmlDoc.CreateElement("Root");
            
            // Ghi du lieu Bt
            XmlElement bt = xmlDoc.CreateElement("BacTru");
            int n = m1.Bt.Length;
            for (int i = 0; i < n;i++ )
            {
                XmlElement bti = xmlDoc.CreateElement("Bt");
                bti.InnerText = m1.Bt[i].ToString();
                bt.AppendChild(bti);
            }
            rootname.AppendChild(bt);

            // Ghi du lieu Be day la ton cap i
            XmlElement bedaylaton = xmlDoc.CreateElement("BeDayLaTon");
            n = m1.RealDeltai.Length;
            for (int i = 0; i < n; i++)
            {
                XmlElement bedaylaton_i = xmlDoc.CreateElement("BeDayLaTon_i");
                bedaylaton_i.InnerText = m1.RealDeltai[i].ToString();
                bedaylaton.AppendChild(bedaylaton_i);
            }
            rootname.AppendChild(bedaylaton);

            // Ghi du lieu Bg
            XmlElement bg = xmlDoc.CreateElement("BacGong");
            n = m1.Bg.Length;
            for (int i = 0; i < n; i++)
            {
                XmlElement bgi = xmlDoc.CreateElement("Bg");
                bgi.InnerText = m1.Bg[i].ToString();
                bg.AppendChild(bgi);
            }
            rootname.AppendChild(bg);

            // Ghi du lieu Hbg
            XmlElement hbg = xmlDoc.CreateElement("HaBacGong");
            n = m1.Hbg.Length;
            for (int i = 0; i < n; i++)
            {
                XmlElement hbgi = xmlDoc.CreateElement("Hbg");
                hbgi.InnerText = m1.Hbg[i].ToString();
                hbg.AppendChild(hbgi);
            }
            rootname.AppendChild(hbg);

            // Ghi du lieu N
            XmlElement nLaton = xmlDoc.CreateElement("SoLaTon");
            n = m1.N.Length;
            for (int i = 0; i < n; i++)
            {
                XmlElement nLatoni = xmlDoc.CreateElement("N");
                nLatoni.InnerText = m1.N[i].ToString();
                nLaton.AppendChild(nLatoni);
            }
            rootname.AppendChild(nLaton);

            // Ghi du lieu Td
            XmlElement nThongDau = xmlDoc.CreateElement("ThongDau");
            n = m1.Thongdaucapton.Length;
            for (int i = 0; i < n; i++)
            {
                XmlElement nThongDaui = xmlDoc.CreateElement("Td");
                nThongDaui.InnerText = m1.Thongdaucapton[i].ToString();
                nThongDau.AppendChild(nThongDaui);
            }
            rootname.AppendChild(nThongDau);

            // Ghi du lieu duong kinh
            XmlElement diameter = xmlDoc.CreateElement("Diameter");
            diameter.InnerText = m1.Diameter.ToString();
            rootname.AppendChild(diameter);

            // Ghi du lieu Ccs
            XmlElement ccs = xmlDoc.CreateElement("Ccs");
            ccs.InnerText = m1.Ccs.ToString();
            rootname.AppendChild(ccs);

            // Ghi du lieu Hcs
            XmlElement hcs = xmlDoc.CreateElement("Hcs");
            hcs.InnerText = m1.Hcs.ToString();
            rootname.AppendChild(hcs);

            // Ghi du lieu Hr
            XmlElement hr = xmlDoc.CreateElement("Hr");
            hr.InnerText = m1.Hr.ToString();
            rootname.AppendChild(hr);

            // Ghi du lieu Delta
            XmlElement delta = xmlDoc.CreateElement("Delta");
            delta.InnerText = m1.Delta.ToString();
            rootname.AppendChild(delta);

            // Ghi du lieu Ke
            XmlElement ke = xmlDoc.CreateElement("Ke");
            ke.InnerText = m1.Ke.ToString();
            rootname.AppendChild(ke);

            // Ghi du lieu lam Tron
            XmlElement lt = xmlDoc.CreateElement("Lt");
            lt.InnerText = m1.Lt.ToString();
            rootname.AppendChild(lt);

            // Ghi du lieu Ks
            XmlElement ks = xmlDoc.CreateElement("Ks");
            ks.InnerText = m1.Ks.ToString();
            rootname.AppendChild(ks);
            // Ghi du lieu CheckBox MultiMaterial
            XmlElement MultiMaterial = xmlDoc.CreateElement("MultiMaterial");
            MultiMaterial.InnerText = m1.IsMultiMaterial.ToString();
            rootname.AppendChild(MultiMaterial);

            // Ghi du lieu CheckBox Giam Chieu dai rau ngang o cap tru khac cap gong
            XmlElement chkGiamChieuDaiRau = xmlDoc.CreateElement("GiamChieuDaiRau");
            chkGiamChieuDaiRau.InnerText = m1.IsGiamChieuDaiRau.ToString();
            rootname.AppendChild(chkGiamChieuDaiRau);

            //Ghi dữ liệu Po
            XmlElement po = xmlDoc.CreateElement("Po");
            po.InnerText = m1.Po.ToString();
            rootname.AppendChild(po);

            //Ghi dữ liệu Uv
            XmlElement uv = xmlDoc.CreateElement("Uv");
            uv.InnerText = m1.Uv.ToString();
            rootname.AppendChild(uv);

            //Ghi dữ liệu Poyc
            XmlElement poyc = xmlDoc.CreateElement("Poyc");
            poyc.InnerText = m1.Poyc.ToString();
            rootname.AppendChild(poyc);

            //Ghi du lieu duong kinh trong
            XmlElement d1t = xmlDoc.CreateElement("D1t");
            XmlElement d2t = xmlDoc.CreateElement("D2t");
            XmlElement d3t = xmlDoc.CreateElement("D3t");
            XmlElement d4t = xmlDoc.CreateElement("D4t");
            XmlElement d5t = xmlDoc.CreateElement("D5t");
            XmlElement d1n = xmlDoc.CreateElement("D1n");
            XmlElement d2n = xmlDoc.CreateElement("D2n");
            XmlElement d3n = xmlDoc.CreateElement("D3n");
            XmlElement d4n = xmlDoc.CreateElement("D4n");
            XmlElement d5n = xmlDoc.CreateElement("D5n");
            XmlElement h1 = xmlDoc.CreateElement("H1");
            XmlElement h2 = xmlDoc.CreateElement("H2");
            XmlElement h3 = xmlDoc.CreateElement("H3");
            XmlElement h4 = xmlDoc.CreateElement("H4");
            XmlElement h5 = xmlDoc.CreateElement("H5");
            XmlElement dd1 = xmlDoc.CreateElement("DemDau1");
            XmlElement dd2 = xmlDoc.CreateElement("DemDau2");
            XmlElement dd3 = xmlDoc.CreateElement("DemDau3");
            XmlElement dd4 = xmlDoc.CreateElement("DemDau4");
            XmlElement dd5 = xmlDoc.CreateElement("DemDau5");
            XmlElement l01 = xmlDoc.CreateElement("L01");
            XmlElement l02 = xmlDoc.CreateElement("L02");
            XmlElement nCanDoc = xmlDoc.CreateElement("SoCanDoc");

      
           

            d1t.InnerText = m1.Winding1.D1t.ToString();
            d2t.InnerText = m1.Winding1.D2t.ToString();
            d3t.InnerText = m1.Winding1.D3t.ToString();
            d4t.InnerText = m1.Winding1.D4t.ToString();
            d5t.InnerText = m1.Winding1.D5t.ToString();

            d1n.InnerText = m1.Winding1.D1n.ToString();
            d2n.InnerText = m1.Winding1.D2n.ToString();
            d3n.InnerText = m1.Winding1.D3n.ToString();
            d4n.InnerText = m1.Winding1.D4n.ToString();
            d5n.InnerText = m1.Winding1.D5n.ToString();

            h1.InnerText = m1.Winding1.H1.ToString();
            h2.InnerText = m1.Winding1.H2.ToString();
            h3.InnerText = m1.Winding1.H3.ToString();
            h4.InnerText = m1.Winding1.H4.ToString();
            h5.InnerText = m1.Winding1.H5.ToString();

            dd1.InnerText = m1.Winding1.Demdau1.ToString();
            dd2.InnerText = m1.Winding1.Demdau2.ToString();
            dd3.InnerText = m1.Winding1.Demdau3.ToString();
            dd4.InnerText = m1.Winding1.Demdau4.ToString();
            dd5.InnerText = m1.Winding1.Demdau5.ToString();

            l01.InnerText = m1.Winding1.L01.ToString();
            l02.InnerText = m1.Winding1.L02.ToString();
            nCanDoc.InnerText = m1.Winding1.NCanDoc.ToString();

            rootname.AppendChild(d1t);
            rootname.AppendChild(d2t);
            rootname.AppendChild(d3t);
            rootname.AppendChild(d4t);
            rootname.AppendChild(d5t);
            rootname.AppendChild(d1n);
            rootname.AppendChild(d2n);
            rootname.AppendChild(d3n);
            rootname.AppendChild(d4n);
            rootname.AppendChild(d5n);
            rootname.AppendChild(h1);
            rootname.AppendChild(h2);
            rootname.AppendChild(h3);
            rootname.AppendChild(h4);
            rootname.AppendChild(h5);

            rootname.AppendChild(dd1);
            rootname.AppendChild(dd2);
            rootname.AppendChild(dd3);
            rootname.AppendChild(dd4);
            rootname.AppendChild(dd5);

            rootname.AppendChild(l01);
            rootname.AppendChild(l02);
            rootname.AppendChild(nCanDoc);

            // Insert xml into doc
            xmlDoc.AppendChild(rootname);
            
            // Save file
            xmlDoc.Save(path);
        }
        public static MBA1 ReadDataFormXml(string path)
        {
            MBA1 m1 = new MBA1();
            XmlDocument docXML = new XmlDocument();
            docXML.Load(path);
            XmlElement root = docXML.DocumentElement;
            //Đọc bậc trụ
            string xpath = "//Bt";
            XmlNodeList lstBt = root.SelectNodes(xpath);
            int vol = lstBt.Count;
            for (int i = 0; i < vol;i++ )
            {
                m1.Bt[i] = double.Parse(lstBt.Item(i).InnerText);
            }

            //Đọc bậc gông
            xpath = "//Bg";
            XmlNodeList lstBg = root.SelectNodes(xpath);
            vol = lstBg.Count;
            for (int i = 0; i < vol; i++)
            {
                m1.Bg[i] = double.Parse(lstBg.Item(i).InnerText);
            }

            //Đọc hạ bậc gông
            xpath = "//Hbg";
            XmlNodeList lstHbg = root.SelectNodes(xpath);
            vol = lstHbg.Count;
            for (int i = 0; i < vol; i++)
            {
                m1.Hbg[i] = double.Parse(lstHbg.Item(i).InnerText);
            }

            //Đọc số lá tôn
            xpath = "//N";
            XmlNodeList lstN = root.SelectNodes(xpath);
            vol = lstN.Count;
            for (int i = 0; i < vol; i++)
            {
                m1.N[i] = int.Parse(lstN.Item(i).InnerText);
            }

           

            //Đọc số vị trí thông dầu
            xpath = "//Td";
            XmlNodeList lstTd = root.SelectNodes(xpath);
            vol = lstTd.Count;
            for (int i = 0; i < vol; i++)
            {
                m1.Thongdaucapton[i] = int.Parse(lstTd.Item(i).InnerText);
            }

            //Đọc Ccs
            xpath = "//Ccs";
            XmlNode nodeCcs = root.SelectSingleNode(xpath);
            m1.Ccs = double.Parse(nodeCcs.InnerText);

            //Đọc cờ Multimateria
            xpath = "//MultiMaterial";
            XmlNode multiMaterial = root.SelectSingleNode(xpath);
            if (multiMaterial!= null) m1.IsMultiMaterial = bool.Parse(multiMaterial.InnerText); else m1.IsMultiMaterial = false;

           //Đọc cờ Giam chieu dai Rau
            xpath = "//GiamChieuDaiRau";
            XmlNode giamChieuDaiRau = root.SelectSingleNode(xpath);
            if (giamChieuDaiRau!= null) m1.IsGiamChieuDaiRau = bool.Parse(giamChieuDaiRau.InnerText); else m1.IsGiamChieuDaiRau = false;

            //Đọc Hs
            xpath = "//Hcs";
            XmlNode nodeHcs = root.SelectSingleNode(xpath);
            m1.Hcs = double.Parse(nodeHcs.InnerText);

            //Đọc đường kính
            xpath = "//Diameter";
            XmlNode nodeDia = root.SelectSingleNode(xpath);
            m1.Diameter = double.Parse(nodeDia.InnerText);

            //Đọc Hr
            xpath = "//Hr";
            XmlNode nodeHr = root.SelectSingleNode(xpath);
            m1.Hr = double.Parse(nodeHr.InnerText);

            //Đọc Delta
            xpath = "//Delta";
            XmlNode nodeDelta = root.SelectSingleNode(xpath);
            m1.Delta = double.Parse(nodeDelta.InnerText);

            //Đọc Ke
            xpath = "//Ke";
            XmlNode nodeKe = root.SelectSingleNode(xpath);
            m1.Ke = double.Parse(nodeKe.InnerText);

            //Đọc be day la ton moi cap
            xpath = "//BeDayLaTon_i";
            XmlNodeList lstBeDayLaTon = root.SelectNodes(xpath);
            vol = lstBeDayLaTon.Count;
            for (int i = 0; i < vol; i++)
            {
                m1.RealDeltai[i] = Double.Parse(lstBeDayLaTon.Item(i).InnerText);
            }

            //Đọc Lt
            xpath = "//Lt";
            XmlNode nodeLt = root.SelectSingleNode(xpath);
            m1.Lt = int.Parse(nodeLt.InnerText);

            //Đọc Ks
            xpath = "//Ks";
            XmlNode nodeKs = root.SelectSingleNode(xpath);
            m1.Ks = double.Parse(nodeKs.InnerText);

            //Đọc Uv:
            xpath = "//Uv";
            XmlNode nodeUv = root.SelectSingleNode(xpath);
            m1.Uv = double.Parse(nodeUv.InnerText);

            //Đọc Po:
            xpath = "//Po";
            XmlNode nodePo = root.SelectSingleNode(xpath);
            m1.Po = double.Parse(nodePo.InnerText);
            //Đọc Po:
            xpath = "//Poyc";
            XmlNode nodePoyc = root.SelectSingleNode(xpath);
            m1.Poyc = double.Parse(nodePoyc.InnerText);

            //Doc boi day
            //Đọc Duong kinh trong
            xpath = "//D1t";
            XmlNode nodeD1t = root.SelectSingleNode(xpath);
            m1.Winding1.D1t = double.Parse(nodeD1t.InnerText);
            xpath = "//D2t";
            XmlNode nodeD2t = root.SelectSingleNode(xpath);
            m1.Winding1.D2t = double.Parse(nodeD2t.InnerText);
            xpath = "//D3t";
            XmlNode nodeD3t = root.SelectSingleNode(xpath);
            m1.Winding1.D3t = double.Parse(nodeD3t.InnerText);
            xpath = "//D4t";
            XmlNode nodeD4t = root.SelectSingleNode(xpath);
            m1.Winding1.D4t = double.Parse(nodeD4t.InnerText);
            xpath = "//D5t";
            XmlNode nodeD5t = root.SelectSingleNode(xpath);
            m1.Winding1.D5t = double.Parse(nodeD5t.InnerText);

            //Đọc Duong kinh ngoai
            xpath = "//D1n";
            XmlNode nodeD1n = root.SelectSingleNode(xpath);
            m1.Winding1.D1n = double.Parse(nodeD1n.InnerText);
            xpath = "//D2n";
            XmlNode nodeD2n = root.SelectSingleNode(xpath);
            m1.Winding1.D2n = double.Parse(nodeD2n.InnerText);
            xpath = "//D3n";
            XmlNode nodeD3n = root.SelectSingleNode(xpath);
            m1.Winding1.D3n = double.Parse(nodeD3n.InnerText);
            xpath = "//D4n";
            XmlNode nodeD4n = root.SelectSingleNode(xpath);
            m1.Winding1.D4n = double.Parse(nodeD4n.InnerText);
            xpath = "//D5n";
            XmlNode nodeD5n = root.SelectSingleNode(xpath);
            m1.Winding1.D5n = double.Parse(nodeD5n.InnerText);

            //Đọc Chiều cao bối dây
            xpath = "//H1";
            XmlNode nodeH1 = root.SelectSingleNode(xpath);
            m1.Winding1.H1 = double.Parse(nodeH1.InnerText);

            xpath = "//H2";
            XmlNode nodeH2 = root.SelectSingleNode(xpath);
            m1.Winding1.H2 = double.Parse(nodeH2.InnerText);

            xpath = "//H3";
            XmlNode nodeH3 = root.SelectSingleNode(xpath);
            m1.Winding1.H3 = double.Parse(nodeH3.InnerText);

            xpath = "//H4";
            XmlNode nodeH4 = root.SelectSingleNode(xpath);
            m1.Winding1.H4 = double.Parse(nodeH4.InnerText);

            xpath = "//H5";
            XmlNode nodeH5 = root.SelectSingleNode(xpath);
            m1.Winding1.H5 = double.Parse(nodeH5.InnerText);

            //Đọc L01,L02
            xpath = "//L01";
            XmlNode nodeL01 = root.SelectSingleNode(xpath);
            m1.Winding1.L01 = double.Parse(nodeL01.InnerText);

            XmlNode nodeL02 = root.SelectSingleNode(xpath);
            m1.Winding1.L02 = double.Parse(nodeL02.InnerText);

            //Đọc số căn dọc
            xpath = "//SoCanDoc";
            XmlNode nodeNCanDoc = root.SelectSingleNode(xpath);
            m1.Winding1.NCanDoc = int.Parse(nodeNCanDoc.InnerText);

            //Đọc đệm đầu
            xpath = "//DemDau1";
            XmlNode nodeDD1 = root.SelectSingleNode(xpath);
            m1.Winding1.Demdau1 = double.Parse(nodeDD1.InnerText);

            xpath = "//DemDau2";
            XmlNode nodeDD2 = root.SelectSingleNode(xpath);
            m1.Winding1.Demdau2 = double.Parse(nodeDD2.InnerText);

            xpath = "//DemDau3";
            XmlNode nodeDD3 = root.SelectSingleNode(xpath);
            m1.Winding1.Demdau3 = double.Parse(nodeDD3.InnerText);

            xpath = "//DemDau4";
            XmlNode nodeDD4 = root.SelectSingleNode(xpath);
            m1.Winding1.Demdau4 = double.Parse(nodeDD4.InnerText);

            xpath = "//DemDau5";
            XmlNode nodeDD5 = root.SelectSingleNode(xpath);
            m1.Winding1.Demdau5 = double.Parse(nodeDD5.InnerText);
            return m1;
        }
    }
}
