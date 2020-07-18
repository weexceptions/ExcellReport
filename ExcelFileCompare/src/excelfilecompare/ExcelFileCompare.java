/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelfilecompare;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author S_All
 */
public class ExcelFileCompare {

    /**
     * @param args the command line arguments
     */
    public static ArrayList getList() {
        ArrayList al = new ArrayList();
        try {
            File file = new File("D:\\Aadhaar_Transacted.xlsx");   //creating a new file instance  
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
            while (itr.hasNext()) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
                            String s = Double.toString(cell.getNumericCellValue());
                            if (s.endsWith("E11")) {
                                s = s.replace("E11", "");
                                s = s.replace(".", "");

//                                System.out.println(s);
                            } else {
//                                System.out.print(cell.getNumericCellValue() + "\t\t\t");
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                            String a = cell.getStringCellValue();
                            if ((a.length() == 12) && (a.startsWith("2200"))) {
//                                System.out.println(a);
                                al.add(a.trim().toString());
                            } else {
//                                System.out.print(a + "\t\t\t");
                            }
                            break;

                        default:
                    }
                }
//                System.out.println("");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return al;
    }

    public static void main(String[] args) {
        ArrayList dist = getList();
/*
        System.out.println("********************************");
        for (Iterator it = dist.iterator(); it.hasNext();) {
            String object = (String) it.next();
            System.out.println(object);
        }
        System.out.println("********************************");
        System.exit(0);*/
        try {
            int count = 0;
            File file = new File("D:\\RationVitranReport.xlsx");   //creating a new file instance  
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
            aa:
            while (itr.hasNext()) {

                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    Cell cell2 = cell;
                    cell.setCellType(CellType.STRING);
//                    System.out.println("ffffffffffffff"+cell2.getStringCellValue());

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
                            String s = Double.toString(cell.getNumericCellValue());
                            System.out.println("=======" + s);
                            System.exit(0);
                            if (s.endsWith("E11")) {
                                System.out.println(s);
                                s = s.replace("E11", "");
                                s = s.replace(".", "");
                                while (s.length() < 12) {
                                    s = s + "0";
                                }
                                if (dist.contains(s)) {
                                    System.out.println("C" + s);
                                    continue aa;
                                } else {
                                    System.out.print(s + "\t");
                                }
                            } else {
                                System.out.print(cell.getNumericCellValue() + "\t");
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                            String a = cell.getStringCellValue().trim().replace(" ", "");
                           
                            if ((a.trim().length() == 12) && (a.trim().startsWith("2200"))) {
                                if ((dist.contains(a.trim()))) {

                                    continue aa;
                                }
                                else{
//                                    System.out.println("oooooooooo");
                                System.out.print("" + a + "\t");}
                            } else {
                                System.out.print(a + "\t");
                            }
                            break;

                        default:
                    }
                }
                System.out.println("");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static ArrayList getOneTimeList() {
        ArrayList al = new ArrayList();
        try {
            File file = new File("D:\\Aadhaar_Transacted.xlsx");   //creating a new file instance  
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
            while (itr.hasNext()) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
                            String s = Double.toString(cell.getNumericCellValue());
                            if (s.endsWith("E11")) {
                                s = s.replace("E11", "");
                                s = s.replace(".", "");

//                                System.out.println(s);
                            } else {
//                                System.out.print(cell.getNumericCellValue() + "\t\t\t");
                            }
                            break;
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                            String a = cell.getStringCellValue();
                            if ((a.length() == 12) && (a.startsWith("2200"))) {
//                                System.out.println(a);
                                al.add(a);
                            } else {
//                                System.out.print(a + "\t\t\t");
                            }
                            break;

                        default:
                    }
                }
//                System.out.println("");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return al;
    }
}
