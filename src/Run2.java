import java.io.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
/* IMPORTANT NOTE: 
 * Each run configuration is customized to the excel file being worked on and the tasks that
 * are to be performed. Input necessarily includes the file location, and information regarding
 * column index for a task to be performed. For example, the column number to be sorted is required
 * for the class SortExcel to work.
 * 
 * Any line with '//<---' after shows where manual input occurred
 */
public class Run2 {
  public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
    //the file being edited
    File file = new File("/Users/MarkPreschern/Desktop/CA White Space.xlsx"); //<--- file name

    Data data;
    Tab tab;

    //parses the file
    ParseExcel pe = new ParseExcel(file);
    tab = pe.getTab();
    //sorts
    SortExcel se = new SortExcel(tab, 2, ""); //<--- column to perform task on
    tab = se.getTab();
    //separates the file into numerous file by unique Territory: Team names
    SeparateExcel sepE = new SeparateExcel(tab, 2, "", false); //<--- column to perform task on
    data = sepE.getData();

    @SuppressWarnings("unused")
    CreateExcel ce = new CreateExcel(data, file, 2, -1); //<--- column to use for naming convention
    System.out.println("Terminated");
  }
}