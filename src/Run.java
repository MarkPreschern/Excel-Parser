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
public class Run {
  public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
    //the file being edited
    File file = new File("/Users/MarkPreschern/Desktop/SCARenewals.xlsx"); //<--- file name

    Data data;
    Tab tab;

    //parses the file
    ParseExcel pe = new ParseExcel(file);
    
    tab = pe.getTab();
    //removes duplicates by Account Name
    RemoveDuplicatesExcel rde = new RemoveDuplicatesExcel(tab, 1, ""); //<--- column to perform task on
    rde.run();
    tab = rde.getTab();
    //sorts the file by Territory: Team
    SortExcel se = new SortExcel(tab, 5, ""); //<--- column to perform task on
    se.run();
    tab = se.getTab();
    //separates the file into numerous files by unique Territory: Team names
    SeparateExcel sepE = new SeparateExcel(tab, 5, "", false); //<--- column to perform task on
    sepE.run();
    data = sepE.getData();
    //sorts each separate file by Account Owner
    for(int i = 0;i < data.excelFiles.size(); i++) {
      SortExcel se2 = new SortExcel(data.excelFiles.get(i).tabs.get(0), 0, "");
      se2.run();
      data.excelFiles.get(i).tabs.get(0).setTab(se2.getTab());
    }
    //sorts each separate file by Fiscal Period
    for(int i = 0;i < data.excelFiles.size(); i++) {
      SortExcel se2 = new SortExcel(data.excelFiles.get(i).tabs.get(0), 9, "");
      se2.run();
      data.excelFiles.get(i).tabs.get(0).setTab(se2.getTab());
    }
    //separates each file into tabs by Fiscal Period
    for(int i = 0;i < data.excelFiles.size(); i++) {
      SeparateExcel sepE2 = new SeparateExcel(data.excelFiles.get(i).tabs.get(0), 9, "", true);
      sepE2.run();
      data.excelFiles.get(i).setExcel(sepE2.getExcel());
    }

    @SuppressWarnings("unused")
    CreateExcel ce = new CreateExcel(data, file, 5, 0); //<--- column to use for naming convention
    System.out.println("Terminated");
  }
}