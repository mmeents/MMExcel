using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using C0DEC0RE;
using Excel = Microsoft.Office.Interop.Excel;

namespace MMExcel {

  public enum StartMode { smOpen, smNew } 
  
  public class MMRng {
    public Excel.Range Rng;
    public string CellA = "";
    public string CellB = "";
    public MMWS Owner;
    public MMRng(MMWS aOwner, string sCellA, string sCellB ){ 
      Owner = aOwner;
      CellA = sCellA;
      CellB = sCellB;
      Rng = aOwner.WS.get_Range(CellA, CellB);
    }
    public string Text {get {return (string)Rng.Text;}} 
    public double ColumnWidth { get { return Rng.ColumnWidth; } set { Rng.ColumnWidth = value; } }
    public double RowHeight { get { return Rng.RowHeight;} set{ Rng.RowHeight = value;} }
  }

  public class MMWS { 
    public Excel.Worksheet WS;
    public MMExcel Owner; 
    public MMWS(MMExcel aOwner, Excel.Worksheet aWS ){
      Owner = aOwner;
      WS = aWS;
    }
    private MMRng getRange(string sCelA, string sCelB){ 
      return  new MMRng(this, sCelA, sCelB);
    } 
    public MMRng this [string sCelA, string sCelB] {get { return getRange(sCelA, sCelB);} }
    public string Name { get { return WS.Name; } set{WS.Name = value;} }     
    public void AddPicture(string sFileName, float dLeft, float dTop, float dWidth, float dHeight){ 
      WS.Shapes.AddPicture(sFileName, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, dLeft, dTop, dWidth, dHeight);
    }
  }

  public class MMExcel {
     
    public string FilePathName = "";
    public Excel.Application xlApp = null;
    public Excel.Workbook xlBook = null;
    public object mv = System.Reflection.Missing.Value;
    public List<MMWS> Sheet = null;

    public MMExcel(StartMode aSM, string aFilePathName) {
      FilePathName = aFilePathName;
      Sheet = new List<MMWS>();      
      switch(aSM) {
        case StartMode.smNew:
          New();
          break;
        case StartMode.smOpen:
          Open(FilePathName);
          break;
      }
    }
    public void Close() {
      if(xlBook != null) {
        try {
          xlBook.CheckCompatibility = false;
          xlBook.SaveAs(FilePathName,Excel.XlFileFormat.xlWorkbookDefault,mv,mv,mv,mv,Excel.XlSaveAsAccessMode.xlExclusive,mv,mv,mv,mv,mv);
        } finally {
          xlBook.Close(true,mv,mv);
        }
        if(xlApp != null) {
          xlApp.Quit();
          ReleaseObject(xlApp);
          ReleaseObject(xlBook);
          foreach(MMWS tw in Sheet) {
            ReleaseObject(tw.WS);
          }
          Sheet.Clear();
        }
      }
    }
    public void New() {
      if(xlBook == null) {
        if (xlApp == null){
          xlApp = new Excel.Application();
        }
        xlBook = xlApp.Workbooks.Add(mv);
        xlBook.CheckCompatibility = false;
        
        MMWS aWS = new MMWS(this, xlBook.Worksheets.get_Item(1));        
        Sheet.Add(aWS);
      } else {
        Close();
        New();
      }      
    }
    public void Open(string sFileNamePath) {
      if(xlBook == null) {
        if(xlApp == null) {
          xlApp = new Excel.Application();
        }
        FilePathName = sFileNamePath;
        xlBook = xlApp.Workbooks.Open(FilePathName,true,false,mv,mv,mv,mv,mv,mv,true,mv,mv,mv,mv,mv);
        xlBook.CheckCompatibility = false;
        Int32 iSheetCount = xlBook.Sheets.Count;
        for(Int32 i = 1;i <= iSheetCount;i++) {
          MMWS aWS = new MMWS(this, xlBook.Worksheets.get_Item(i));          
          Sheet.Add(aWS);
        }
      } else {
        Close();
        Open(sFileNamePath);
      }
    }
    public Excel.Worksheet this [Int32 iSheetIndx] {get { return Sheet[iSheetIndx].WS;}}
    public string ReadCellText(Int32 iSheetIndx,Int32 iRow,Int32 iCol) {      
      return (string)(this[iSheetIndx].Cells[iRow,iCol] as Excel.Range).Text;
    }
    public object ReadCellValue(Int32 iSheetIndx, Int32 iRow, Int32 iCol){ 
      return (this[iSheetIndx].Cells[iRow,iCol] as Excel.Range).Value2();
    }

    private void ReleaseObject(object obj) {
      try {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      } catch(Exception ex) {
        obj = null;
      } finally {
        GC.Collect();
      }
    }

  }
}
