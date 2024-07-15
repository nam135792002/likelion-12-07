package vn.edu.likelion.app;

import vn.edu.likelion.entity.DocumentApp;
import vn.edu.likelion.entity.WorkbookApp;

public class App {
    public static void main(String[] args){
        WorkbookApp workbookApp = new WorkbookApp();
        DocumentApp documentApp = new DocumentApp();

        documentApp.writeFileDocument();
        workbookApp.writeFileExcel();

        documentApp.print();
        workbookApp.print();

        workbookApp.writeFileAvailability();

        workbookApp.createNewFile();
    }
}
