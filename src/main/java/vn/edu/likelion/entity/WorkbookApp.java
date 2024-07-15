package vn.edu.likelion.entity;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import vn.edu.likelion.utility.Constant;

import java.io.*;
import java.util.Base64;

public class WorkbookApp {
    private XSSFWorkbook workbook;
    private Sheet sheet;

    public WorkbookApp() {
        this.workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet("Sheet1");
    }

    public void writeFileExcel(){
        String line;
        Row row;
        Cell cell1;
        Cell cell2;

        int i = 0;

        try {
            BufferedReader bufferedReader = new BufferedReader(new FileReader(Constant.SOURCE_FILE));
            while ((line = bufferedReader.readLine()) != null){
                String[] s = line.split("\t");
                if(s[2].equals("1")){
                    String encode = Base64.getEncoder().encodeToString(s[1].getBytes());
                    row = sheet.createRow(i);
                    cell1 = row.createCell(0);
                    cell1.setCellValue(s[0]);

                    cell2 = row.createCell(1);
                    cell2.setCellValue(encode);

                    i++;
                }
            }
        }catch (IOException e){
            System.out.println("Taking database fail!");
        }


        try {
            FileOutputStream stream = new FileOutputStream(Constant.FILE_PRESENT);
            workbook.write(stream);
            workbook.close();
            stream.close();
        } catch (IOException e) {
            System.out.println("Writing file excel fail!");
        }
    }

    public void print(){
        File file = new File(Constant.FILE_PRESENT);

        try {
            FileInputStream inputStream = new FileInputStream(file);
            Workbook workbookXlsx = WorkbookFactory.create(inputStream);
            sheet = workbookXlsx.getSheetAt(0);
            inputStream.close();
            workbookXlsx.close();
        } catch (IOException e) {
            System.out.println("Reading file excel fail!");
        }

        System.out.println("List of student is present: ");
        for (Row row1 : sheet){
            String id = row1.getCell(0).toString();
            String encodeName = row1.getCell(1).toString();

            String decodeName = new String(Base64.getDecoder().decode(encodeName));
            System.out.println(id + "\t" + decodeName);
        }
    }

    public void writeFileAvailability(){
        String line;
        Row row;
        Cell cell;

        File file = new File("output.xlsx");
        Workbook workbookXlsx = null;
        try {
            FileInputStream inputStream = new FileInputStream(file);
            workbookXlsx = WorkbookFactory.create(inputStream);
            sheet = workbookXlsx.getSheetAt(0);
            inputStream.close();
        } catch (IOException e) {
            System.out.println("Reading file excel fail!");
        }

        int i = 4;

        try {
            BufferedReader bufferedReader = new BufferedReader(new FileReader(Constant.SOURCE_FILE));
            while ((line = bufferedReader.readLine()) != null){
                String[] s = line.split("\t");
                if(s[2].equals("1")){
                    row = sheet.getRow(i++);
                    cell = row.getCell(1);
                    cell.setCellValue(s[1]);
                }
            }
        }catch (IOException e){
            System.out.println("Taking database fail!");
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(file);
            workbookXlsx.write(outputStream);
            workbookXlsx.close();
            outputStream.close();
        } catch (IOException e) {
            e.getStackTrace();
        }
    }

    public void createNewFile(){
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("trainee");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Merge cell");

        sheet.addMergedRegion(new CellRangeAddress(0,2,0,1));
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream("NewList.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
