/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelfilecompare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author S_All
 */
public class ExcelFileCompare1 {

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
        Workbook workbook=null;
        try {
            int count = 0;
            workbook = new XSSFWorkbook();
            Sheet sOut = workbook.createSheet("Results");
            File file = new File("D:\\RationVitranReport.xlsx");   //creating a new file instance  
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
            aa:
            while (itr.hasNext()) {

                Row row = itr.next();
                Row row2 = null;
                int cellCOunt = 0;

                    int xAxis = 0;
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    cell.setCellType(CellType.STRING);
//                    System.out.println("ffffffffffffff"+cell2.getStringCellValue());

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
                            String s = Double.toString(cell.getNumericCellValue()); ///NOT NEEDED
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
                                } else {
//                                    System.out.println("oooooooooo");
                                   
                                row2 = sOut.createRow(count);
                                Cell c = row2.createCell(xAxis);
                                c.setCellType(CellType.STRING);
                                c.setCellValue(a);
                                xAxis++;
                                count++;
                                    System.out.print( a + "\t");
                                }
                            } else {
                                 row2.createCell(xAxis).setCellValue(cell.getStringCellValue());
//                                 System.out.println("xa"+xAxis);
//                                    c=cell;
                                    xAxis++;
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
        finally{try{
            FileOutputStream fileOut = new FileOutputStream("D:\\Report.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close(); } catch (Exception e) {
            e.printStackTrace();
        }
        }
    }

    public static ArrayList getOneTimeList() {
        ArrayList al = new ArrayList();
        try {
            File file = new File("D:\\Aadhaar_Transacted.xlsx");   //creating a new file instance  
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook wb = new XSSFWorkbook();
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
