import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;

//sorts the data by the index and then separates
public class SeparateExcel extends Task {
  ArrayList<ArrayList<XSSFRow>> seperatedRows = new ArrayList<ArrayList<XSSFRow>>();
  ArrayList<XSSFRow> rows;
  boolean makeTabs;

  Data data;
  Excel excel;

  //constructor
  public SeparateExcel(Tab tab, int index, String name, boolean makeTabs) {
    super(tab, index, name);
    this.rows = tab.rows;
    this.makeTabs = makeTabs;
  }

  //runs the program
  public void run() {
    this.seperate();
    this.createData();
  }

  //separates rows by uniqueness into individual lists after sorting it
  public void seperate() {
    //sorts data first
    Task se = new SortExcel(new Tab(this.rows), index, "");
    se.run();
    this.rows = se.getTab().rows;

    DataFormatter formatter = new DataFormatter();

    if(this.rows.size() > 1) { 
      XSSFRow header = rows.remove(0);
      XSSFRow next = rows.remove(0);
      String comp = formatter.formatCellValue(next.getCell(index));
      ArrayList<XSSFRow> temp = new ArrayList<XSSFRow>();
      temp.add(header);
      temp.add(next);
      while (this.rows.size() > 0) {
        XSSFRow r = this.rows.get(0);
        if(r.getCell(index) != null) {
          if(r.getCell(index).getStringCellValue().equals(comp)) {
            temp.add(this.rows.remove(0));
          } else {
            comp = formatter.formatCellValue(r.getCell(index));
            this.seperatedRows.add(temp);
            temp = new ArrayList<XSSFRow>();
            temp.add(header);
            temp.add(this.rows.remove(0));
          }
        } else {
          this.rows.remove(0);
        }
      }
      this.seperatedRows.add(temp);
      System.out.println("Seperated");
    }
  }

  //sets variables for either a tab or file separation
  public void createData() {

    //for separate tabs
    ArrayList<Tab> at = new ArrayList<Tab>();
    for (ArrayList<XSSFRow> r : this.seperatedRows) {
      at.add(new Tab(r));
    }
    excel = new Excel(at);

    //for separate files
    ArrayList<Excel> ae = new ArrayList<Excel>();
    for (ArrayList<XSSFRow> r : this.seperatedRows) {
      ae.add(new Excel(new ArrayList<Tab>(Arrays.asList(new Tab(r)))));
    }
    data = new Data(ae);
  }

  //for separate tabs
  public Excel getExcel() {
    return this.excel;
  }

  //for separate files
  public Data getData() {
    return this.data;
  }

  //sets rows to t's rows
  public void setRows(Tab t) {
    this.rows = t.rows;
    seperatedRows = new ArrayList<ArrayList<XSSFRow>>();
  }
}