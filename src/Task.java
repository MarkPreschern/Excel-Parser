//represents a function that is performed on data
public class Task {
  Tab tab;
  int index;
  String name;

  //constructor
  public Task(Tab tab, int index, String name) {
    this.tab = tab;
    this.index = index;
    this.name = name;
  }

  //to run individual tasks
  public void run() {}

  //to get a tab from a task
  public Tab getTab() {
    return null;
  }

  //to get an excel from a task
  public Excel getExcel() {
    return null;
  }

  //to get data from a task
  public Data getData() {
    return null;
  }
  
  //sets rows to r
  public void setRows(Tab t) {}
}
