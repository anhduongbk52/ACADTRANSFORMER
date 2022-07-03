// (C) Copyright 2016 by  
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using ACADTRANSFORMER.GIAODIEN;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(ACADTRANSFORMER.MyCommands))]

namespace ACADTRANSFORMER
{
    public class MyCommands
    {
        [CommandMethod("MTT")]
        public void MyCommand()
        {
            Window1 myWindow = new Window1();
            myWindow.Show();
        }
    }
}
