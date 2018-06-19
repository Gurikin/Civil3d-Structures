// Copyright 2009-2010 by Autodesk, Inc.
//
//Permission to use, copy, modify, and distribute this software in
//object code form for any purpose and without fee is hereby granted, 
//provided that the above copyright notice appears in all copies and 
//that both that copyright notice and the limited warranty and
//restricted rights notice below appear in all supporting 
//documentation.
//
//AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
//AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
//MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
//DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
//UNINTERRUPTED OR ERROR FREE.
//
//Use, duplication, or disclosure by the U.S. Government is subject to 
//restrictions set forth in FAR 52.227-19 (Commercial Computer
//Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
//(Rights in Technical Data and Computer Software), as applicable.

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

    public class StucturePipes : Rt.IExtensionApplication {
        private static Microsoft.Office.Interop.Excel.Application xlApp=null;
        private static Workbook xlWb = null;
        private static Worksheet xlWsStructures=null;
        private static Worksheet xlWsPipes=null;


        [Rt.CommandMethod("StructurePipesLabels")]
        public void StructurePipesLabels() {
            Ed.Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Civ.CivilDocument doc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            // Document AcadDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            // Check that there's a pipe network to parse
            if ( doc.GetPipeNetworkIds() == null ) {
                ed.WriteMessage("There are no pipe networks to export.  Open a document that contains at least one pipe network");
                return;
            }

            
            // Iterate through each pipe network
            using ( Db.Transaction ts = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction() ) {
                
                Dictionary<string, char> dictPipe = new Dictionary<string, char>(); // track data parts column

                Ed.PromptEntityOptions opt = new Ed.PromptEntityOptions("\nSelect an Alignment");
                opt.SetRejectMessage("\nObject must be an alignment.\n");
                opt.AddAllowedClass(typeof(CivDb.Alignment), false);
                Db.ObjectId alignID = ed.GetEntity(opt).ObjectId;

                try {
                    Db.ObjectId oNetworkId = doc.GetPipeNetworkIds()[0];
                    CivDb.Network oNetwork = ts.GetObject(oNetworkId, Db.OpenMode.ForWrite) as CivDb.Network;

                    //// Get pipes:
                    //ObjectIdCollection oPipeIds = oNetwork.GetPipeIds();
                    //int pipeCount = oPipeIds.Count;

                    //// we can edit the slope, so make that column yellow
                    //Range colRange = xlWsPipes.get_Range("D1", "D" + ( pipeCount + 1 ));
                    //colRange.Interior.ColorIndex = 6;
                    //colRange.Interior.Pattern = XlPattern.xlPatternSolid;
                    //colRange.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;

                    //foreach ( ObjectId oid in oPipeIds ) {

                    //    Pipe oPipe = ts.GetObject(oid, OpenMode.ForRead) as Pipe;
                    //    ed.WriteMessage("   " + oPipe.Name);
                    //    col = 'B';
                    //    row++;
                    //    aRange = xlWsPipes.get_Range("A" + row, System.Type.Missing);
                    //    aRange.Value2 = oPipe.Handle.Value.ToString();
                    //    aRange = xlWsPipes.get_Range("B" + row, System.Type.Missing);
                    //    aRange.Value2 = oPipe.StartPoint.X + "," + oPipe.StartPoint.Y + "," + oPipe.StartPoint.Z;
                    //    aRange = xlWsPipes.get_Range("C" + row, System.Type.Missing);
                    //    aRange.Value2 = oPipe.EndPoint.X + "," + oPipe.EndPoint.Y + "," + oPipe.EndPoint.Z;

                    //    // This only gives the absolute value of the slope:
                    //    // aRange = xlWsPipes.get_Range("D" + row, System.Type.Missing);
                    //    // aRange.Value2 = oPipe.Slope;
                    //    // This gives a signed value:
                    //    aRange = xlWsPipes.get_Range("D" + row, System.Type.Missing);
                    //    aRange.Value2 = ( oPipe.EndPoint.Z - oPipe.StartPoint.Z ) / oPipe.Length2DCenterToCenter;

                    //    // Get the catalog data to use later
                    //    ObjectId partsListId = doc.Styles.PartsListSet["Standard"];
                    //    PartsList oPartsList = ts.GetObject(partsListId, OpenMode.ForRead) as PartsList;
                    //    ObjectIdCollection oPipeFamilyIdCollection = oPartsList.GetPartFamilyIdsByDomain(DomainType.Pipe);

                    //    foreach ( PartDataField oPartDataField in oPipe.PartData.GetAllDataFields() ) {
                    //        // Make sure the data has a column in Excel, if not, add the column
                    //        if ( !dictPipe.ContainsKey(oPartDataField.ContextString) ) {
                    //            char nextCol = ( char )( ( int )'E' + dictPipe.Count );
                    //            aRange = xlWsPipes.get_Range("" + nextCol + "1", System.Type.Missing);
                    //            aRange.Value2 = oPartDataField.ContextString + "(" + oPartDataField.Name + ")";
                    //            dictPipe.Add(oPartDataField.ContextString, nextCol);

                    //            // We can edit inner diameter or width, so make those yellow
                    //            if ( ( oPartDataField.ContextString == "PipeInnerDiameter" ) || ( oPartDataField.ContextString == "PipeInnerWidth" ) ) {
                    //                colRange = xlWsPipes.get_Range("" + nextCol + "1", "" + nextCol + ( pipeCount + 1 ));
                    //                colRange.Interior.ColorIndex = 6;
                    //                colRange.Interior.Pattern = XlPattern.xlPatternSolid;
                    //                colRange.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
                    //            }

                    //            // Check the part catalog data to see if this part data is user-modifiable
                    //            foreach ( ObjectId oPipeFamliyId in oPipeFamilyIdCollection ) {
                    //                PartFamily oPartFamily = ts.GetObject(oPipeFamliyId, OpenMode.ForRead) as PartFamily;
                    //                SizeFilterField oSizeFilterField = null;

                    //                try {
                    //                    oSizeFilterField = oPartFamily.PartSizeFilter[oPartDataField.Name];
                    //                } catch ( System.Exception e ) { }

                    //                /* You can also iterate through all defined size filter fields this way:
                    //                 SizeFilterRecord oSizeFilterRecord = oPartFamily.PartSizeFilter;
                    //                 for ( int i = 0; i < oSizeFilterRecord.ParamCount; i++ ) {
                    //                     oSizeFilterField = oSizeFilterRecord[i];
                    //                  } */

                    //                if ( oSizeFilterField != null ) {
                    //                         // Check whether it can be modified:
                    //                         if ( oSizeFilterField.DataSource == PartDataSourceType.Optional ) {
                    //                             colRange = xlWsPipes.get_Range("" + nextCol + "1", "" + nextCol + ( pipeCount + 1 ));
                    //                             colRange.Interior.ColorIndex = 4;
                    //                             colRange.Interior.Pattern = XlPattern.xlPatternSolid;
                    //                             colRange.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
                    //                         }

                    //                         break;
                    //                     }

                    //            }
                    //        }
                    //        char iColumnPipes = dictPipe[oPartDataField.ContextString];
                    //        aRange = aRange = xlWsPipes.get_Range("" + iColumnPipes + row, System.Type.Missing);
                    //        aRange.Value2 = oPartDataField.Value;

                    //    }
                    //}

                    // Now export the structures

                    Dictionary<string, char> dictStructures = new Dictionary<string, char>(); // track data parts column
                    
                    // Get structures:
                    Db.ObjectIdCollection oStructureIds = oNetwork.GetStructureIds();
                    foreach ( Db.ObjectId oid in oStructureIds ) {
                        CivDb.Structure oStructure = ts.GetObject(oid, Db.OpenMode.ForRead) as CivDb.Structure;
                        //col = 'B';
                        //row++;
                        //aRange = xlWsStructures.get_Range("" + col + row, System.Type.Missing);
                        //aRange.Value2 = oStructure.Handle.Value;
                        //aRange = xlWsStructures.get_Range("" + ++col + row, System.Type.Missing);
                        //aRange.Value2 = oStructure.Position.X + "," + oStructure.Position.Y + "," + oStructure.Position.Z;

                        foreach ( CivDb.PartDataField oPartDataField in oStructure.PartData.GetAllDataFields() ) {
                            // Make sure the data has a column in Excel, if not, add the column
                            if ( !dictStructures.ContainsKey(oPartDataField.ContextString) ) {
                                char nextCol = ( char )( ( int )'D' + dictStructures.Count );
                                //aRange = xlWsStructures.get_Range("" + nextCol + "1", System.Type.Missing);
                                //aRange.Value2 = oPartDataField.ContextString;
                                dictStructures.Add(oPartDataField.ContextString, nextCol);

                            }
                            char iColumnStructure = dictStructures[oPartDataField.ContextString];
                            //aRange = aRange = xlWsStructures.get_Range("" + iColumnStructure + row, System.Type.Missing);
                            //aRange.Value2 = oPartDataField.Value;
                        }
                    }

                } catch ( Autodesk.AutoCAD.Runtime.Exception ex ) {
                    ed.WriteMessage("PipeSample: " + ex.Message);
                    return;
                }

            }
        }

        #region IExtensionApplication Members

        public void Initialize () { }

        public void Terminate ()  {
            // Clean up all our Excel COM objects    
            // This will close Excel without saving

            if ( xlWb != null ) {
                try {
                    xlWb.Close(false, Type.Missing, Type.Missing);
                    xlApp.Quit();
                    Marshal.FinalReleaseComObject(xlWsStructures);
                    Marshal.FinalReleaseComObject(xlWsPipes);
                    Marshal.FinalReleaseComObject(xlWb);
                    Marshal.FinalReleaseComObject(xlApp);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                } catch ( System.Exception e ) {
                    Console.WriteLine(e.Message);
                }
            }            
        }

        #endregion
    }
}
