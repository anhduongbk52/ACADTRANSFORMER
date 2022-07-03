using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
namespace ACADTRANSFORMER.THUVIENCAD
{
    //Clas chua cac ham ve 2D trong wpf
    class GrapWPF
    {
        // chen textblock vao canvas
        public static void drawText(string txt, double x, double y, Color color, Canvas canvasObj)
        {
            TextBlock textBlock = new TextBlock();

            textBlock.Text = txt;

            textBlock.Foreground = new SolidColorBrush(color);

            Canvas.SetLeft(textBlock, x);

            Canvas.SetTop(textBlock, y);

            canvasObj.Children.Add(textBlock);
        }

        // chen Hinh vuong vao canvas biet tam canh trai
        public static void drawRecCenterLeft(double x, double y, double width, double height, double bedaynet, Color color, Canvas canvasObj, string txtToolTip)
        {
            System.Windows.Shapes.Rectangle rec = new System.Windows.Shapes.Rectangle
            {
                Stroke = Brushes.Black,
                Fill = new SolidColorBrush(color),
                StrokeThickness = bedaynet,
                Height = height,
                Width = width,
                ToolTip=txtToolTip
            };
            Canvas.SetTop(rec, y-height/2);
            Canvas.SetLeft(rec, x );
            canvasObj.Children.Add(rec);    
        }
        // chen Hinh vuong vao canvas biet tam canh phai
        public static void drawRecCenterRight(double x, double y, double width, double height, double bedaynet, Color color, Canvas canvasObj,string txtToolTip)
        {
            System.Windows.Shapes.Rectangle rec = new System.Windows.Shapes.Rectangle
            {
                Stroke = Brushes.Black,
                Fill = new SolidColorBrush(color),
                StrokeThickness = bedaynet,
                Height = height,
                Width = width,
                ToolTip = txtToolTip
            };
            Canvas.SetTop(rec, y - height / 2);
            Canvas.SetLeft(rec, x-width);
            canvasObj.Children.Add(rec);
        }

    }
}
