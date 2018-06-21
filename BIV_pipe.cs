// Copyright BIV

using System;
using System.Collections.Generic;
using System.Text;
using Rt = Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Civ = Autodesk.Civil.ApplicationServices;
using Ap = Autodesk.AutoCAD.ApplicationServices;
using Ed = Autodesk.AutoCAD.EditorInput;
using Db = Autodesk.AutoCAD.DatabaseServices;
using CivDb = Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System.Runtime.InteropServices;
using System.Collections;
using System.Text.RegularExpressions;

namespace PipeData {
    /// <summary>
    /// This is an example of how to work with pipe networks, and how to 
    /// set up interop to import and export data to and from MS Excel.
    /// 
    /// This sample has two commands:
    /// 
    /// ExportToExcel : iterates through the first pipe network in a drawing,
    /// and dumps data about the component pipes and structures to an Excel spreadsheet.
    /// 
    /// ImportFromExcel : reads the spreadsheet from the ExportToExcel command, and applies
    /// any changes to the pipe network.  This command could be adapted with the 
    /// Workbook.Open() method to read a previously saved spreadsheet.
    /// </summary>

    public class BIVPipe {
        
        public List<Db.ObjectId> GetPipeIdByName(string[] pipeNames) {
            Ed.Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Civ.CivilDocument civDoc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            var pipeIds = new List<Db.ObjectId>();
                        
            // Iterate through each pipe network
            using ( Db.Transaction ts = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction() ) {
                
                try {
                    foreach (Db.ObjectId networkId in civDoc.GetPipeNetworkIds()) {
                        CivDb.Network oNetwork = ts.GetObject(networkId, Db.OpenMode.ForWrite) as CivDb.Network;
                        foreach (Db.ObjectId pipeId in oNetwork.GetPipeIds()) {
                            CivDb.Pipe oPipe = ts.GetObject(pipeId, Db.OpenMode.ForRead) as CivDb.Pipe;
                            foreach (string name in pipeNames) {
                                if (oPipe.Name == name) {
                                    pipeIds.Add(oPipe.Id);
                                }                                
                            }
                        }                        
                    }
                    return pipeIds;                    
                } catch ( Autodesk.AutoCAD.Runtime.Exception ex ) {
                    ed.WriteMessage("StructurePipesData: " + ex.Message);
                    return null;
                }
            }
        }
    }
}
