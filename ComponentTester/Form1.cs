using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MMExcel;
using Excel = Microsoft.Office.Interop.Excel;
using C0DEC0RE;

namespace ComponentTester
{
  public partial class Form1:Form
  {
    public Form1() {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e) {
      
      string sFilePathName = EnsureDestFileUnique(MMExt.UserLogLocation() + "Excel01.xlsx");
      MMExcel.MMExcel mm = new MMExcel.MMExcel(StartMode.smNew, sFilePathName);
      MMWS ws0 = mm.Sheet[0];
      ws0.Name = "ZeroSheet";
      double dInchValue = 0.25;         
      ws0["A1", "AB120"].RowHeight = dInchValue.toPointsVertical();  
      ws0["A1", "AB120"].ColumnWidth =  dInchValue.toPointsHorizontal(); // dColWidthPerPoint * dInchValue.toPointsHorizontal();  
      ws0["A1", "A1"].Rng.Font.Name = "Century Gothic";
      Excel.Font fontA = ws0["A1", "A1"].Rng.Font;
      


           
      mm.Close();

      Process.Start(sFilePathName);

    }


    public string EnsureDestFileUnique(string DestFileName) {
      string sFN = DestFileName;
      if(File.Exists(DestFileName)) {  // ensure file with name does not exist.        
        Int32 iCounter = 0;
        String sPath00 = Path.GetDirectoryName(DestFileName) + @"\";
        String sFileN = Path.GetFileNameWithoutExtension(DestFileName);
        String sExt = Path.GetExtension(DestFileName);
        while(File.Exists(sFN)) {
          iCounter++;          
          sFN = sPath00 + sFileN + "x" + Convert.ToString(iCounter) + sExt;
        }  
      } 
      return sFN;
    }
  }
}
