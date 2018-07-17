import java.awt.Color;
import java.io.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//writes a new excel file given a single or list of data 
public class CreateExcel {
  Data data;
  File file;
  int index;
  int colorIndex;

  //constructor for a single excel file
  public CreateExcel(Data data, File file, int index, int colorIndex) 
      throws IOException, InvalidFormatException {
    this.data = data;
    this.file = file;
    this.index = index;
    this.colorIndex = colorIndex;
    this.createFiles();
  }

  //creates multiple files
  public void createFiles() throws IOException, InvalidFormatException {
    //creates a zip for the excel files to be stored
    File zipFile = new File(FilenameUtils.removeExtension(file.getPath()) + " - Collection.zip");
    //File zipFile = new File("/Users/MarkPreschern/Desktop/Collection.zip");
    ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile));
    System.out.println("Your zip file has been generated!");
    //creates excel files

    for(Excel e : this.data.excelFiles) {
      this.createFile(e.tabs, zos);
    }
    zos.close();
  }

  //creates a single file
  public void createFile(ArrayList<Tab> tabs, ZipOutputStream zos) 
      throws IOException, InvalidFormatException {
    XSSFWorkbook workbook = new XSSFWorkbook();
    DataFormatter formatter = new DataFormatter();
    String excelFileName = "";

    //creates each tab as a sheet
    for(Tab tab : tabs) {
      ArrayList<XSSFRow> rows = tab.rows;
      if (rows.size() > 1) {    
        XSSFSheet sheet = workbook.createSheet();

        //creates file name
        excelFileName = formatter.formatCellValue(rows.get(1).getCell(index)) + ".xlsx";

        //adjusts cell style for the first row
        XSSFCellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        boldStyle.setBorderBottom(BorderStyle.valueOf((short)3));
        boldStyle.setBorderTop(BorderStyle.valueOf((short)2));
        boldStyle.setBorderRight(BorderStyle.valueOf((short)2));
        boldStyle.setBorderLeft(BorderStyle.valueOf((short)2));
        XSSFFont boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldStyle.setFont(boldFont);

        int colorPlaceholder = 0;
        String colorStringValue = "";

        //places cell data from rows onto workbook
        for (int i = 0; i < rows.size(); i++) {
          XSSFRow row = sheet.createRow(i);
          XSSFRow thisRow = rows.get(i);
          for (int j = 0; j < thisRow.getPhysicalNumberOfCells(); j++) {        
            XSSFCell cell = thisRow.getCell(j);  
            String val = formatter.formatCellValue(cell);
            
            XSSFCell c = row.createCell(j);
            c.setCellValue(val);

            //creates the regular cell style
            XSSFCellStyle regularStyle = workbook.createCellStyle();
            regularStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
            regularStyle.setBorderBottom(BorderStyle.valueOf((short)2));
            regularStyle.setBorderTop(BorderStyle.valueOf((short)2));
            regularStyle.setBorderRight(BorderStyle.valueOf((short)2));
            regularStyle.setBorderLeft(BorderStyle.valueOf((short)2));

            //creates color separation in file
            if (this.colorIndex != -1 && i != 0) {
              String comparison = formatter.formatCellValue(rows.get(i).getCell(this.colorIndex));
              if (!comparison.equals(colorStringValue)) {
                colorPlaceholder ++;
                colorStringValue = comparison;
              }
              //sets color separation in file
              if (colorPlaceholder % 2 == 1) {
                regularStyle.setFillForegroundColor(new XSSFColor(new Color(204, 255, 255)));
              } else {
                regularStyle.setFillForegroundColor(new XSSFColor(new Color(255, 204, 153)));
              }
              regularStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            //sets cell style
            if (i == 0) {
              c.setCellStyle(boldStyle);
            } else {
              c.setCellStyle(regularStyle);
            }

            //auto-sizes columns
            if (i == rows.size() - 1) {
              sheet.autoSizeColumn(j);
            }
          }
        }
        sheet.createFreezePane(0, 1);
      }
    }

    //writes the excel file to the given zip file
    if (!excelFileName.equals("")) {
      //puts excel file into the zip
      zos.putNextEntry(new ZipEntry(
          FilenameUtils.removeExtension(file.getName()) + " - " + excelFileName));
      ByteArrayOutputStream bos = new ByteArrayOutputStream();
      //writes the excel file to the zip
      workbook.write(bos);
      bos.writeTo(zos);
      zos.closeEntry();
      workbook.close();
      System.out.println("Your excel file has been generated!");
    }
  }
}