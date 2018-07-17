import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFRow;

//contains multiple excel files
class Data {
  ArrayList<Excel> excelFiles;
  
  //constructor
  public Data(ArrayList<Excel> excelFiles) {
    this.excelFiles = excelFiles;
  }
  
  //sets the rows in tab to t
  public void setData(Data d) {
    this.excelFiles = d.excelFiles;
  }
}

//contains data for a single excel file
class Excel {
  ArrayList<Tab> tabs;
  
  //constructor
  public Excel(ArrayList<Tab> tabs) {
    this.tabs = tabs;
  }
  
  //sets the tabs in excel to e
  public void setExcel(Excel e) {
    this.tabs = e.tabs;
  }
}

//contains information for all cells in a tab
class Tab {
  ArrayList<XSSFRow> rows;
  
  //constructor
  public Tab(ArrayList<XSSFRow> rows) {
    this.rows = rows;
  }
  
  //sets the rows in tab to t
  public void setTab(Tab t) {
    this.rows = t.rows;
  }
}
