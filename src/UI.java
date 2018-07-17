import java.io.*;
import java.util.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFRow;

//allows the user to put in input to parse and perform tasks on a file
public class UI {
  public static void main(String[] args) 
      throws FileNotFoundException, InvalidFormatException, IOException {

    //gets a file
    Scanner scan = new Scanner(System.in);
    File file = getFile(scan);

    //parses the given file
    ParseExcel pe = new ParseExcel(file);
    Tab tab = pe.getTab();

    //finds the max cells in a tab
    int maxCells = getMaxCells(tab);

    //holds all tasks
    ArrayList<Task> tasks = createTasks(scan, tab, maxCells);

    //performs the tasks and returns the updated data
    Data data = performTasks(tab, tasks);

    //creates the excel file
    createFiles(scan, data, file, tab, maxCells);
  }

  //asks the user to enter a valid file, and parses it
  public static File getFile(Scanner scan) 
      throws FileNotFoundException, InvalidFormatException, IOException {
    File file = new File("");

    //asks the user to provide a file location to be parsed
    boolean correctFile = false;
    while (correctFile == false) {
      System.out.println("\nExcel Parser\n");
      System.out.println("Type in/Paste the file location of the file you would like to parse. "
          + "If the file location/name is incorrect, you will be prompted to paste it again.");

      File f = new File(scan.nextLine());
      if(f.exists() && !f.isDirectory()) { 
        file = f;
        correctFile = true;
      }
    }
    return file;
  }

  public static ArrayList<Task> createTasks(Scanner scan, Tab tab, int maxCells) {
    //holds tasks the user wants to be performed
    ArrayList<Task> tasks = new ArrayList<Task>();
    String input = "";
    while(!input.equals("pt")) {
      System.out.println("\nTask Manager\n");
      System.out.println("Type in the letters shown to add a task, show your task list, or "
          + "perform the tasks.");

      System.out.println("'rd' : remove duplicates and null values.");
      System.out.println("'sort' : sort values in alphabetical order.");
      System.out.println("'s1' : separates unique values into different excel files.");
      System.out.println("'s2' : separates unique values into different excel tabs.");
      System.out.println("'st' : see the tasks you've added so far and/or delete tasks.");
      System.out.println("'pt' : perform the tasks.");

      input = scan.nextLine();

      //column index for the task is provided by the user below
      int index = -1;

      if (input.equals("rd") || input.equals("sort") || input.equals("s1") || input.equals("s2")) {
        //asks the user for a column number to perform the task on
        while (index == -1) {
          System.out.println("Column Number for Task: (0 - " + (maxCells - 1) + ")");
          int temp = scan.nextInt();
          if (temp >= 0 && temp < maxCells) {
            index = temp;
          }
        }
      }

      //creates the tasks
      if (input.equals("rd")) {
        Task t = new RemoveDuplicatesExcel(tab, index, "Remove Duplicates");
        tasks.add(t);
      } else if (input.equals("sort")) {
        Task t = new SortExcel(tab, index, "Sort");
        tasks.add(t);
      } else if (input.equals("s1")) {
        Task t = new SeparateExcel(tab, index, "Separate Files", false);
        tasks.add(t);
      } else if (input.equals("s2")) {
        Task t = new SeparateExcel(tab, index,"Separate Tabs", true);
        tasks.add(t);
      } else if (input.equals("st")) {
        tasks = displayTasks(scan, tasks);
      }
    }
    return tasks;
  }

  //displays tasks added and allows the user to delete tasks
  public static ArrayList<Task> displayTasks(Scanner scan, ArrayList<Task> tasks) {
    int input = -2;
    while (input != -1) {
      System.out.println("Tasks Created\n");
      for (int i = 0;i < tasks.size();i++) {
        System.out.println((i + 1) + ": " + tasks.get(i).name + ", Column " + tasks.get(i).index);
      }
      System.out.println("\nType the number next to a task to delete it. Otherwise, type '-1' to return"
          + " to the task manager.");

      input = scan.nextInt();  
      //removes a task if it is within the correct bounds and a task exists
      if (tasks.size() > 0 && input > 0 && input <= tasks.size()) {
        tasks.remove(input - 1);
      }
    }
    return tasks;
  }

  //performs all tasks on the tab
  public static Data performTasks(Tab tab, ArrayList<Task> tasks) {
    //creates a data object containing the tab
    Data data = new Data(new ArrayList<Excel>(Arrays.asList(
        new Excel(new ArrayList<Tab>(Arrays.asList(tab))))));

    //performs each task in order on data
    for (Task task : tasks) {
      Data tempData;
      if (task.name.equals("Separate Files") || task.name.equals("Separate Tabs")) {
        tempData = new Data(new ArrayList<Excel>(Arrays.asList(
        new Excel(new ArrayList<Tab>(Arrays.asList(
            new Tab(new ArrayList<XSSFRow>())))))));
      } else {
        tempData = data;
      }
      
      //manipulates data depending on the task
      for (int i = 0;i < data.excelFiles.size();i++) {
        Excel e = data.excelFiles.get(i);
        for (int j = 0;j < e.tabs.size();j++) {
          Tab t = e.tabs.get(j);

          task.setRows(t);
          task.run();
          if (task.name.equals("Separate Files")) {
            tempData.excelFiles.addAll(task.getData().excelFiles);
          } else if (task.name.equals("Separate Tabs")) {
            tempData.excelFiles.add(task.getExcel());
          } else {
            tempData.excelFiles.get(i).tabs.get(j).setTab(task.getTab());
          }
        }
      }
      data.setData(tempData);
    }
    return data;
  }

  //creates the file with the user's color preferance
  @SuppressWarnings("unused")
  public static void createFiles(Scanner scan, Data data, File file, Tab tab, int maxCells) 
      throws InvalidFormatException, IOException {
    String line = "";
    int colorIndex = -1;
    int namingIndex = -1;

    //sets name index and color index
    while (true) {
      System.out.println("\nColumn Number for Naming Convention: (0 - " + (maxCells - 1) + ")");
      int tempIndex = scan.nextInt();
      if (tempIndex >= 0 && tempIndex < maxCells) {
        namingIndex = tempIndex;
        break;
      }
    }
    
    while(!(line.equals("t") || line.equals("f"))) {
      System.out.println("\nType 't' for color separation by unique column values or 'f' otherwise.");
      line = scan.nextLine();
      if (line.equals("t")) {
        inner: while (true) {
          System.out.println("\nColumn Number for Color Separation: (0 - " + (maxCells - 1) + ")");
          int tempIndex = scan.nextInt();
          if (tempIndex >= 0 && tempIndex < maxCells) {
            colorIndex = tempIndex;
            break inner;
          }
        }
      }
    }

    CreateExcel ce = new CreateExcel(data, file, namingIndex, colorIndex);
    System.out.println("Terminated");
  }

  //finds the max columns in a tab
  public static int getMaxCells(Tab tab) {
    int maxCells = 0;
    for (XSSFRow r: tab.rows) {
      if (r.getPhysicalNumberOfCells() > maxCells) {
        maxCells = r.getPhysicalNumberOfCells();
      }
    }
    return maxCells;
  }
}