import java.util.ArrayList;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;

//removes duplicates and null values from the list
public class RemoveDuplicatesExcel extends Task {
  ArrayList<XSSFRow> rows;

  //constructor
  public RemoveDuplicatesExcel(Tab tab, int index, String name) {
    super(tab, index, name);
    this.rows = tab.rows;
  }

  //runs the program
  public void run() {
    this.remove(new ArrayList<XSSFRow>());
  }

  //removes duplicates and null cells  
  public void remove(ArrayList<XSSFRow> acc) {
    DataFormatter formatter = new DataFormatter();
    XSSFRow r1 = this.rows.remove(0);
    String s1 = formatter.formatCellValue(r1.getCell(index));
    if (s1 != null && s1 != "") {
      boolean isEqual = false;
      for(XSSFRow r : acc) {
        String s2 = formatter.formatCellValue(r.getCell(index));
        if (s1.equals(s2)) {
          isEqual = true;
        }
      }
      if (isEqual == false) {
        acc.add(r1);
      }
    }

    //recurs if there are still elements in this.rows
    if (this.rows.size() > 0) {
      this.remove(acc);
    } else {
      this.rows = acc;
      System.out.println("Removed Duplicates");
    }
  }

  //returns a list without duplicates
  public Tab getTab() {
    return new Tab(this.rows);
  }

  //sets rows to t's rows
  public void setRows(Tab t) {
    this.rows = t.rows;
  }
}