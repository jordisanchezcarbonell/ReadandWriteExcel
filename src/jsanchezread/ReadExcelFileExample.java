package jsanchezread;

import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;
import jsanchezread.Student;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Created by anirudh on 20/10/14.
 */
public class ReadExcelFileExample {

    private static final String FILE_PATH = "C:\\Users\\jordi\\Desktop\\TEST\\testReadStudents.xlsx";
    private static final String FILE_PATH_sinRep = "C:\\Users\\jordi\\Desktop\\TEST\\testReadStudentssinrep.xlsx";
    public static void main(String args[]) {

        List<Student> studentList = getStudentsListFromExcel();
            Set setA = new HashSet();
        System.out.println(studentList);
        System.out.println("------------------------------");
        
        //eliminamos los duplicados
                setA.addAll(studentList);

        System.out.println(setA);
  
        writeStudentsListToExcel(setA);
       
    }

    private static List<Student> getStudentsListFromExcel() {
        List<Student> studentList = new ArrayList<Student>();
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(FILE_PATH);

            // Using XSSF for xlsx format, for xls use HSSF
            Workbook workbook = new XSSFWorkbook(fis);

            int numberOfSheets = workbook.getNumberOfSheets();

            //looping over each workbook sheet
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                Iterator<Row> rowIterator = sheet.iterator();

                //iterating over each row
                while (rowIterator.hasNext()) {

                    Student student = new Student();
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    //Iterating over each cell (column wise)  in a particular row.
                    while (cellIterator.hasNext()) {

                        Cell cell = cellIterator.next();
                        //The Cell Containing String will is name.
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                student.setName(cell.getStringCellValue());
                                
                                //The Cell Containing numeric value will contain marks
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                //Cell with index 1 contains marks in Maths
                                if (cell.getColumnIndex() == 1) {
                                    student.setMaths(String.valueOf(cell.getNumericCellValue()));
                                }
                                //Cell with index 2 contains marks in Science
                                else if (cell.getColumnIndex() == 2) {
                                    student.setScience(String.valueOf(cell.getNumericCellValue()));
                                }
                                //Cell with index 3 contains marks in English
                                else if (cell.getColumnIndex() == 3) {
                                    student.setEnglish(String.valueOf(cell.getNumericCellValue()));
                                }   break;
                            case Cell.CELL_TYPE_BLANK:
                                
                                break;
                            default:
                                break;
                        }
                    }
                    //end iterating a row, add all the elements of a row in list
                    if(student.getName()!=null){
                        studentList.add(student);
                    }
                }
            }

            fis.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return studentList;
    }

   public static void writeStudentsListToExcel(Set<Student> studentList){

        // Using XSSF for xlsx format, for xls use HSSF
        Workbook workbook = new XSSFWorkbook();

        Sheet studentsSheet = workbook.createSheet("Students");

        int rowIndex = 0;
        for(Student student : studentList){
            Row row = studentsSheet.createRow(rowIndex++);
            int cellIndex = 0;
            
            if(row != null){
             if(student.getName()!=null){   
            //firstplace in row is name
            row.createCell(cellIndex++).setCellValue(student.getName());

            //second place in row is marks in maths
            row.createCell(cellIndex++).setCellValue(student.getMaths());

            //third place in row is marks in Science
            row.createCell(cellIndex++).setCellValue(student.getScience());

            //fourth place in row is marks in English
            row.createCell(cellIndex++).setCellValue(student.getEnglish());
             }
        }
        }
        //write this workbook in excel file.
        try {
            FileOutputStream fos = new FileOutputStream(FILE_PATH_sinRep);
            workbook.write(fos);
            fos.close();

            System.out.println(FILE_PATH_sinRep + " is successfully written");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }


}