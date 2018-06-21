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

    public class StucturePipes : Rt.IExtensionApplication {
        
        [Rt.CommandMethod("StructurePipesLabels")]
        public void StructurePipesLabels() {
            Ed.Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Civ.CivilDocument civDoc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            // Document AcadDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            // Check that there's a pipe network to parse
            if ( civDoc.GetPipeNetworkIds() == null ) {
                ed.WriteMessage("There are no pipe networks to export.  Open a document that contains at least one pipe network");
                return;
            }

            
            // Iterate through each pipe network
            using ( Db.Transaction ts = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction() ) {
                
                Dictionary<string, char> dictPipe = new Dictionary<string, char>(); // track data parts column

                Ed.PromptEntityOptions opt = new Ed.PromptEntityOptions("\nSelect an Structure");
                opt.SetRejectMessage("\nObject must be an Structure.\n");
                opt.AddAllowedClass(typeof(CivDb.Structure), false);
                Db.ObjectId structureID = ed.GetEntity(opt).ObjectId;

                try {
                    
                    //////Db.ObjectId oNetworkId = civDoc.GetPipeNetworkIds()[0];
                    //////CivDb.Network oNetwork = ts.GetObject(oNetworkId, Db.OpenMode.ForWrite) as CivDb.Network;

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

                    //Db.ObjectIdCollection oStructureIds = oNetwork.GetStructureIds();
                    //foreach ( Db.ObjectId oid in structureID ) {
                    CivDb.Structure oStructure = ts.GetObject(structureID, Db.OpenMode.ForRead) as CivDb.Structure;
                    CivDb.Network oNetwork = ts.GetObject(oStructure.NetworkId, Db.OpenMode.ForWrite) as CivDb.Network;
                    string[] connPipeNames = oStructure.GetConnectedPipeNames();
                    PipeData.BIVPipe pipe = new PipeData.BIVPipe();
                    List<Db.ObjectId> pipeIds = pipe.GetPipeIdByName(connPipeNames);
                    ed.WriteMessage("\nК колодцу присоеденены следующие трубы: ");
                    foreach (Db.ObjectId pipeId in pipeIds) {
                        CivDb.Pipe oPipe = ts.GetObject(pipeId, Db.OpenMode.ForRead) as CivDb.Pipe;
                        ed.WriteMessage("{0} | {1:0.000}   |||   ", oPipe.Name, oStructure.get_PipeCenterDepth(new int()));                        
                    }
                        //col = 'B';
                        //row++;
                        //aRange = xlWsStructures.get_Range("" + col + row, System.Type.Missing);
                        //aRange.Value2 = oStructure.Handle.Value;
                        //aRange = xlWsStructures.get_Range("" + ++col + row, System.Type.Missing);
                        //aRange.Value2 = oStructure.Position.X + "," + oStructure.Position.Y + "," + oStructure.Position.Z;

                        //foreach ( CivDb.PartDataField oPartDataField in oStructure.PartData.GetAllDataFields() ) {
                        //    // Make sure the data has a column in Excel, if not, add the column
                        //    if ( !dictStructures.ContainsKey(oPartDataField.ContextString) ) {
                        //        char nextCol = ( char )( ( int )'D' + dictStructures.Count );
                        //        //aRange = xlWsStructures.get_Range("" + nextCol + "1", System.Type.Missing);
                        //        //aRange.Value2 = oPartDataField.ContextString;
                        //        dictStructures.Add(oPartDataField.ContextString, nextCol);

                        //    }
                        //    char iColumnStructure = dictStructures[oPartDataField.ContextString];
                        //ed.WriteMessage("\npartDataField.Name: " + oPartDataField.Name + "   ===   ColumnStructure to string: " + iColumnStructure + "   ===   PartDataField.Value: " + oPartDataField.Value.ToString() + "\n");
                        //    //aRange = aRange = xlWsStructures.get_Range("" + iColumnStructure + row, System.Type.Missing);
                        //    //aRange.Value2 = oPartDataField.Value;
                        //}
                    //}

                } catch ( Autodesk.AutoCAD.Runtime.Exception ex ) {
                    ed.WriteMessage("StructurePipesData: " + ex.Message);
                    return;
                }
            }
        }

        #region IExtensionApplication Members

        public void Initialize () { }

        public void Terminate ()  {
            // Clean up all our Excel COM objects    
            // This will close Excel without saving

            try {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            } catch ( System.Exception e ) {
                Console.WriteLine(e.Message);
            }
        }
        #endregion
    }
}
