import java.io.*;
import java.util.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//parses the excel file
public class ParseExcel {
  File file;
  ArrayList<XSSFRow> rows = new ArrayList<XSSFRow>();

  //constructor
  public ParseExcel(File file) throws FileNotFoundException, IOException, InvalidFormatException {
    this.file = file;
    this.parse();
  }

  //parses through the excel file creating an ArrayList of rows
  public void parse() throws FileNotFoundException, IOException, InvalidFormatException {
    OPCPackage fs = OPCPackage.open(file);
    @SuppressWarnings("resource")
    XSSFWorkbook wb = new XSSFWorkbook(fs);
    XSSFSheet sheet = wb.getSheetAt(0);

    int rows; // No of rows
    rows = sheet.getPhysicalNumberOfRows();

    for(int r = 0; r < rows; r++) {
      XSSFRow row = sheet.getRow(r);
      if(row != null) {
        this.rows.add(row);
      }
    }
    System.out.println("Parsed");
  }

  //returns one list of rows
  public Tab getTab() {
    return new Tab(this.rows);
  }
}