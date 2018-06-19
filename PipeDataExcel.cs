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
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Civil.DatabaseServices;
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

    public class PipeExcel : IExtensionApplication {
        private static Microsoft.Office.Interop.Excel.Application xlApp=null;
        private static Workbook xlWb = null;
        private static Worksheet xlWsStructures=null;
        private static Worksheet xlWsPipes=null;


        [CommandMethod("ExportToExcel")]
        public void ExportToExcel () {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            CivilDocument doc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            // Document AcadDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            // Check that there's a pipe network to parse
            if ( doc.GetPipeNetworkIds() == null ) {
                ed.WriteMessage("There are no pipe networks to export.  Open a document that contains at least one pipe network");
                return;
            }

            // Interop code is adapted from the MSDN site:
            // http://msdn.microsoft.com/en-us/library/ms173186(VS.80).aspx
            xlApp = new Microsoft.Office.Interop.Excel.Application();

            if ( xlApp == null ) {
                Console.WriteLine(@"EXCEL could not be started. Check that your office installation 
                and project references are correct.");
                return;
            }
            xlApp.Visible = true;
            xlWb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            xlWsStructures = ( Worksheet )xlWb.Worksheets[1];
            xlWsPipes = ( Worksheet )xlWb.Worksheets.Add(xlWsStructures, System.Type.Missing, 1, System.Type.Missing);
            xlWsPipes.Name = "Pipes";
            xlWsStructures.Name = "Structures";

            if ( xlWsPipes == null ) {
                Console.WriteLine(@"Worksheet could not be created. Check that your office installation 
                and project references are correct.");
            }

            // Iterate through each pipe network
            using ( Transaction ts = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction() ) {

                int row = 1;
                char col = 'B';
                Dictionary<string, char> dictPipe = new Dictionary<string, char>(); // track data parts column
                // set up header
                Range aRange = xlWsPipes.get_Range("A1", System.Type.Missing);
                aRange.Value2 = "Handle";
                aRange = xlWsPipes.get_Range("B1", System.Type.Missing);
                aRange.Value2 = "Start";
                aRange = xlWsPipes.get_Range("C1", System.Type.Missing);
                aRange.Value2 = "End";
                aRange = xlWsPipes.get_Range("D1", System.Type.Missing);
                aRange.Value2 = "Slope";
                

                try {
                    ObjectId oNetworkId = doc.GetPipeNetworkIds()[0];
                    Network oNetwork = ts.GetObject(oNetworkId, OpenMode.ForWrite) as Network;

                    // Get pipes:
                    ObjectIdCollection oPipeIds = oNetwork.GetPipeIds();
                    int pipeCount = oPipeIds.Count;

                    // we can edit the slope, so make that column yellow
                    Range colRange = xlWsPipes.get_Range("D1", "D" + ( pipeCount + 1 ));
                    colRange.Interior.ColorIndex = 6;
                    colRange.Interior.Pattern = XlPattern.xlPatternSolid;
                    colRange.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;

                    foreach ( ObjectId oid in oPipeIds ) {

                        Pipe oPipe = ts.GetObject(oid, OpenMode.ForRead) as Pipe;
                        ed.WriteMessage("   " + oPipe.Name);
                        col = 'B';
                        row++;
                        aRange = xlWsPipes.get_Range("A" + row, System.Type.Missing);
                        aRange.Value2 = oPipe.Handle.Value.ToString();
                        aRange = xlWsPipes.get_Range("B" + row, System.Type.Missing);
                        aRange.Value2 = oPipe.StartPoint.X + "," + oPipe.StartPoint.Y + "," + oPipe.StartPoint.Z;
                        aRange = xlWsPipes.get_Range("C" + row, System.Type.Missing);
                        aRange.Value2 = oPipe.EndPoint.X + "," + oPipe.EndPoint.Y + "," + oPipe.EndPoint.Z;
                        
                        // This only gives the absolute value of the slope:
                        // aRange = xlWsPipes.get_Range("D" + row, System.Type.Missing);
                        // aRange.Value2 = oPipe.Slope;
                        // This gives a signed value:
                        aRange = xlWsPipes.get_Range("D" + row, System.Type.Missing);
                        aRange.Value2 = ( oPipe.EndPoint.Z - oPipe.StartPoint.Z ) / oPipe.Length2DCenterToCenter;

                        // Get the catalog data to use later
                        ObjectId partsListId = doc.Styles.PartsListSet["Standard"];
                        PartsList oPartsList = ts.GetObject(partsListId, OpenMode.ForRead) as PartsList;
                        ObjectIdCollection oPipeFamilyIdCollection = oPartsList.GetPartFamilyIdsByDomain(DomainType.Pipe);

                        foreach ( PartDataField oPartDataField in oPipe.PartData.GetAllDataFields() ) {
                            // Make sure the data has a column in Excel, if not, add the column
                            if ( !dictPipe.ContainsKey(oPartDataField.ContextString) ) {
                                char nextCol = ( char )( ( int )'E' + dictPipe.Count );
                                aRange = xlWsPipes.get_Range("" + nextCol + "1", System.Type.Missing);
                                aRange.Value2 = oPartDataField.ContextString + "(" + oPartDataField.Name + ")";
                                dictPipe.Add(oPartDataField.ContextString, nextCol);

                                // We can edit inner diameter or width, so make those yellow
                                if ( ( oPartDataField.ContextString == "PipeInnerDiameter" ) || ( oPartDataField.ContextString == "PipeInnerWidth" ) ) {
                                    colRange = xlWsPipes.get_Range("" + nextCol + "1", "" + nextCol + ( pipeCount + 1 ));
                                    colRange.Interior.ColorIndex = 6;
                                    colRange.Interior.Pattern = XlPattern.xlPatternSolid;
                                    colRange.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
                                }

                                // Check the part catalog data to see if this part data is user-modifiable
                                foreach ( ObjectId oPipeFamliyId in oPipeFamilyIdCollection ) {
                                    PartFamily oPartFamily = ts.GetObject(oPipeFamliyId, OpenMode.ForRead) as PartFamily;
                                    SizeFilterField oSizeFilterField = null;
                                    
                                    try {
                                        oSizeFilterField = oPartFamily.PartSizeFilter[oPartDataField.Name];
                                    } catch ( System.Exception e ) { }

                                    /* You can also iterate through all defined size filter fields this way:
                                     SizeFilterRecord oSizeFilterRecord = oPartFamily.PartSizeFilter;
                                     for ( int i = 0; i < oSizeFilterRecord.ParamCount; i++ ) {
                                         oSizeFilterField = oSizeFilterRecord[i];
                                      } */

                                    if ( oSizeFilterField != null ) {
                                             // Check whether it can be modified:
                                             if ( oSizeFilterField.DataSource == PartDataSourceType.Optional ) {
                                                 colRange = xlWsPipes.get_Range("" + nextCol + "1", "" + nextCol + ( pipeCount + 1 ));
                                                 colRange.Interior.ColorIndex = 4;
                                                 colRange.Interior.Pattern = XlPattern.xlPatternSolid;
                                                 colRange.Interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
                                             }

                                             break;
                                         }
                                    
                                }
                            }
                            char iColumnPipes = dictPipe[oPartDataField.ContextString];
                            aRange = aRange = xlWsPipes.get_Range("" + iColumnPipes + row, System.Type.Missing);
                            aRange.Value2 = oPartDataField.Value;
                            
                        }
                    }

                    // Now export the structures
                    row = 1;
                    col = 'B';
                    Dictionary<string, char> dictStructures = new Dictionary<string, char>(); // track data parts column
                    // Set up header
                    aRange = xlWsStructures.get_Range("A1", System.Type.Missing);
                    aRange.Value2 = "Handle";
                    aRange = xlWsStructures.get_Range("B1", System.Type.Missing);
                    aRange.Value2 = "Location";

                    // Get structures:
                    ObjectIdCollection oStructureIds = oNetwork.GetStructureIds();
                    foreach ( ObjectId oid in oStructureIds ) {
                        Structure oStructure = ts.GetObject(oid, OpenMode.ForRead) as Structure;
                        col = 'B';
                        row++;
                        aRange = xlWsStructures.get_Range("" + col + row, System.Type.Missing);
                        aRange.Value2 = oStructure.Handle.Value;
                        aRange = xlWsStructures.get_Range("" + ++col + row, System.Type.Missing);
                        aRange.Value2 = oStructure.Position.X + "," + oStructure.Position.Y + "," + oStructure.Position.Z;

                        foreach ( PartDataField oPartDataField in oStructure.PartData.GetAllDataFields() ) {
                            // Make sure the data has a column in Excel, if not, add the column
                            if ( !dictStructures.ContainsKey(oPartDataField.ContextString) ) {
                                char nextCol = ( char )( ( int )'D' + dictStructures.Count );
                                aRange = xlWsStructures.get_Range("" + nextCol + "1", System.Type.Missing);
                                aRange.Value2 = oPartDataField.ContextString;
                                dictStructures.Add(oPartDataField.ContextString, nextCol);

                            }
                            char iColumnStructure = dictStructures[oPartDataField.ContextString];
                            aRange = aRange = xlWsStructures.get_Range("" + iColumnStructure + row, System.Type.Missing);
                            aRange.Value2 = oPartDataField.Value;
                        }
                    }

                } catch ( Autodesk.AutoCAD.Runtime.Exception ex ) {
                    ed.WriteMessage("PipeSample: " + ex.Message);
                    return;
                }

            }
        }


        [CommandMethod("ImportFromExcel")]
        public void ImportFromExcel () {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            CivilDocument doc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Document acadDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            // Check that there's a pipe network to parse
            if ( doc.GetPipeNetworkIds() == null ) {
                ed.WriteMessage(@"There are no pipe networks to export. 
                Open a document that contains at least one pipe network");
                return;
            }

            // Interop code is adapted from the MSDN site:
            // http://msdn.microsoft.com/en-us/library/ms173186(VS.80).aspx

            if ( xlApp == null ) {
                ed.WriteMessage("No current Excel spreadsheet.  Run the ExportToExcel command first.");
                return;
            }

            xlApp.Visible = true;
            Workbook wb = xlApp.ActiveWorkbook;
            Worksheet ws = ( Worksheet )wb.Worksheets["Pipes"];
            Dictionary<string, char> dictPipes = new Dictionary<string, char>();
            char col = 'A';
            int row = 1;

            Range aRange = ws.get_Range("" + col + row, System.Type.Missing);
            while ( aRange.Value2 != null ) {
                dictPipes.Add(( String )aRange.Value2, col);
                col++;
                aRange = ws.get_Range("" + col + row, System.Type.Missing);
            }

            foreach ( KeyValuePair<string, char> kvp in dictPipes ) {
                ed.WriteMessage(kvp.Value + " : " + kvp.Key + "\n");
            }

            col = 'A';
            row++;

            using ( Transaction ts = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction() ) {
                aRange = ws.get_Range("" + col + row, System.Type.Missing);
                Handle oHandle;
                while ( aRange.Value2 != null ) {

                    oHandle = new Handle(Int64.Parse(aRange.Value2.ToString()));
                    ObjectId oAcadObjectId = acadDoc.Database.GetObjectId(false, oHandle, 0);
                    Object oAcadObject = ts.GetObject(oAcadObjectId, OpenMode.ForWrite);
                    Pipe oPipe = null;
                    if ( oAcadObject.GetType() == typeof(Pipe) ) {
                        oPipe = ( Pipe )oAcadObject;
                        ed.WriteMessage("Pipe: " + oPipe.Name + " (" + oHandle.Value + "\n");
                    } else {
                        // next loop
                        row++;
                        aRange = ws.get_Range("" + col + row, System.Type.Missing);
                        continue;
                    }

                    // get the shape for the pipe
                    col = dictPipes["SweptShape(CSS)"];
                    aRange = ws.get_Range("" + col + row, System.Type.Missing);
                    string shape = ( String )aRange.Value2;
                    // only works with round and egg shaped:
                    double dia;
                    if ( shape.Contains("EggShaped") ) {
                        // egg shaped based on inner width
                        col = dictPipes["PipeInnerWidth(PIW)"];
                        aRange = ws.get_Range("" + col + row, System.Type.Missing);
                        dia = ( double )aRange.Value2;
                    } else {
                        // round sizes based on diameter:
                        col = dictPipes["PipeInnerDiameter(PID)"];
                        aRange = ws.get_Range("" + col + row, System.Type.Missing);
                        dia = ( double )aRange.Value2;
                    }
                     

                    // resize the pipe:
                    oPipe.ResizeByInnerDiameterOrWidth(dia, true);

                    // adjust the slope
                    col = dictPipes["Slope"];
                    aRange = ws.get_Range("" + col + row, System.Type.Missing);
                    double slope = ( double )aRange.Value2;
                    oPipe.SetSlopeHoldStart(slope);

                    
                    // disconnect and reconnect the pipe to the structure to force the structure to update for the new slope
                    ObjectId oStructureId = oPipe.StartStructureId;
                    oPipe.Disconnect(ConnectorPositionType.Start);
                    oPipe.ConnectToStructure(ConnectorPositionType.Start, oStructureId, true);
                    oStructureId = oPipe.EndStructureId;
                    oPipe.Disconnect(ConnectorPositionType.End);
                    oPipe.ConnectToStructure(ConnectorPositionType.End, oStructureId, true);

                    // Look over the spreadsheet for runtime values (green, index = 4)
                    PartDataRecord record = oPipe.PartData;
                    foreach ( KeyValuePair<string, char> kvp in dictPipes ) {
                        aRange = ws.get_Range("" + kvp.Value + row, System.Type.Missing);
                        if ( ( int )aRange.Interior.ColorIndex == 4 && aRange.Value2 != null ) {
                            // This is a runtime value
                            String context = kvp.Key;
                            // has format: context string (name), use regex to get name:
                            Match m = Regex.Match(context, @"\((\w+)\)");
                            context = m.Groups[1].ToString();
                            
                            PartDataField oPartDataField = record.GetDataFieldBy(context);

                            // todo: append null field?

                            oPartDataField.Value = aRange.Value2;
                            oPipe.PartData = record;                           
                        }
                    }

                    row++;
                    col = 'A';
                    aRange = ws.get_Range("" + col + row, System.Type.Missing);
                }
                ts.Commit();
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
