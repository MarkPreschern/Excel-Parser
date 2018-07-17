import java.util.*;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;

//sorts the list in alphabetical order
public class SortExcel extends Task{
  ArrayList<XSSFRow> rows;

  //constructor
  public SortExcel(Tab tab, int index, String name) {
    super(tab, index, name);
    this.rows = tab.rows;
  }

  //runs the program
  public void run() {
    this.sort();
  }

  //sorts the list
  public void sort() {
    DataFormatter formatter = new DataFormatter();
    Collections.sort(this.rows, new Comparator<XSSFRow>() {
      @Override
      public int compare(XSSFRow o1, XSSFRow o2) {
        if (o1.getCell(index) == null) {
          return -1;
        } else if (o2.getCell(index) == null) {
          return 1;
        } else if (o1.getCell(index).getCellStyle().getFont().getBold()) {
          return -1;
        } else if (o2.getCell(index).getCellStyle().getFont().getBold()) {
          return 1;
        } else {
          String val1 = formatter.formatCellValue(o1.getCell(index));
          String val2 = formatter.formatCellValue(o2.getCell(index));
          return val1.compareTo(val2);
        }
      }
    }); 
    System.out.println("Sorted");
  }

  //returns one list of sorted rows
  public Tab getTab() {
    return new Tab(this.rows);
  }

  //sets rows to t's rows
  public void setRows(Tab t) {
    this.rows = t.rows;
  }
}