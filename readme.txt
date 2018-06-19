//
// (C) Copyright 2010 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//
///////////////////////////////////////////////////////////////////////////////

This samples illustrates how to work with pipe networks and move data to and
from MS Excel using COM interop.

It defines two commands:
ExportToExcel - exports pipe data to excel
ImportFromExcel - imports the same pipe data from excel, updating the pipe network

Building the sample: This project requires two Microsoft Office interop
libraries (Microsoft Office 12.0 Object Library, and Excel 12.0 Object Library)
installed with Excel, as well as AecBaseMgd, AecDBMgd, AcMgd, and AeccDbMgd.

Using the sample: Open a drawing with at least one pipe network and run the ExportToExcel command.  
This exports information about the pipe network pipes and structures to an Excel document.  
Change values as required in the spreadsheet, then run the ImportFromExcel command to apply the 
changed values to the open document.



