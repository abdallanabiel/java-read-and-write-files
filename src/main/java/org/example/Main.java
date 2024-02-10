package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;



public class Main {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "./data/Data Sheet.xlsx";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Valid Data");
            int rowNum = 0;  // number row in new sheet for static titles
            int rowNum2 = 1; //start row number  where data will be input

            //access sheet 1 and 2  to get data
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook.getSheetAt(1);

            Row newRow = newSheet.createRow(rowNum); // make a row object in new sheet and give it the number of fixed row


            int cellNum = 0;//start cell in static row
            //give titles for static row using func setCellValue
            newRow.createCell(cellNum++).setCellValue("studentId");
            newRow.createCell(cellNum++).setCellValue("studentFirstName1");
            newRow.createCell(cellNum++).setCellValue("studentFamilyName1");
            newRow.createCell(cellNum++).setCellValue("studentThirdName1");
            newRow.createCell(cellNum++).setCellValue("studentFourthName1");
            newRow.createCell(cellNum++).setCellValue("studentBirthCountry1");
            newRow.createCell(cellNum++).setCellValue("studentNationality1");

            newRow.createCell(cellNum++).setCellValue("gender");
            newRow.createCell(cellNum++).setCellValue("studentReligion1");
            newRow.createCell(cellNum++).setCellValue("studentNationalityCategory1");
            newRow.createCell(cellNum).setCellValue("schoolId");

            for (Row row : sheet1) {

        //    access sheet 1 and get data
                Cell studentIdCell = row.getCell(0);

                Cell studentFirstName = row.getCell(1);
                Cell studentFamilyName = row.getCell(2);
                Cell studentThirdName = row.getCell(3);
                Cell studentFourthName = row.getCell(4);
                Cell studentBirthCountry = row.getCell(5);
                Cell studentNationality = row.getCell(6);
                Cell genderCell = row.getCell(7);
                Cell studentReligion = row.getCell(8);
                Cell studentNationalityCategory = row.getCell(9);
                Cell schoolIdCell = row.getCell(10);
                //convert data of cells to string using func I have created below " getCellValue()" to avoid any error
                String studentId = getCellValue(studentIdCell);

                String studentFamilyName1 = getCellValue(studentFamilyName);
                String studentFirstName1 = getCellValue(studentFirstName);
                String studentThirdName1 = getCellValue(studentThirdName);
                String studentFourthName1 = getCellValue(studentFourthName);
                String studentBirthCountry1 = getCellValue(studentBirthCountry);
                String studentNationality1 = getCellValue(studentNationality);
                String gender = getCellValue(genderCell);
                String studentReligion1 = getCellValue(studentReligion);
                String studentNationalityCategory1 = getCellValue(studentNationalityCategory);
                    String schoolId = getCellValue(schoolIdCell);



                    for (Row row2 : sheet2) {
                        Cell schoolIdCell2 = row2.getCell(0); // Assuming School ID is in the 1st column (index 0)
                        Cell genderCell2 = row2.getCell(2); // Assuming Gender is in the 3rd column (index 2)


                            String schoolId2 = getCellValue(schoolIdCell2);
                            String gender2 = getCellValue(genderCell2);


                          //conditions on data
                            if (schoolId.equals(schoolId2)&& studentId.length()==10 &&studentId.startsWith("1234") &&( gender.equals(gender2) ||gender2.equals("Mixed") )){

                             /*   System.out.println("Student Details: " +
                                     "ID"           +" "+ studentId +" , "+
                                        "First Name: " + studentFirstName1 + ", " +
                                        "Family Name: " + studentFamilyName1 + ", " +
                                        "Third Name: " + studentThirdName1 + ", " +
                                        "Fourth Name: " + studentFourthName1 + ", " +
                                        "Birth Country: " + studentBirthCountry1+ ", " +
                                        "Nationality: " + studentNationality1 + ", " +
                                        "Gender: " + gender + ", " +
                                        "Religion: " + studentReligion1 + ", " +
                                        "Nationality Category: " + studentNationalityCategory1 + ", " +
                                        "School ID: " + schoolId);*/





                              //insert data in new sheet below static row
                                Row newRow2 = newSheet.createRow(rowNum2++);
                                int cellNum2 = 0;
                                newRow2.createCell(cellNum2++).setCellValue(studentId);
                                newRow2.createCell(cellNum2++).setCellValue(studentFirstName1);
                                newRow2.createCell(cellNum2++).setCellValue(studentFamilyName1);
                                newRow2.createCell(cellNum2++).setCellValue(studentThirdName1);
                                newRow2.createCell(cellNum2++).setCellValue(studentFourthName1);
                                newRow2.createCell(cellNum2++).setCellValue(studentBirthCountry1);
                                newRow2.createCell(cellNum2++).setCellValue(studentNationality1);

                                newRow2.createCell(cellNum2++).setCellValue(gender);
                                newRow2.createCell(cellNum2++).setCellValue(studentReligion1);
                                newRow2.createCell(cellNum2++).setCellValue(studentNationalityCategory1);
                                newRow2.createCell(cellNum2).setCellValue(schoolId);



                            }

                    }

            }
            String resultFilePath = "./data/ValidData.xlsx";
            try (FileOutputStream fos = new FileOutputStream(resultFilePath)) {
                newWorkbook.write(fos);
            } catch (IOException ex) {
                ex.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
// func that convert value of cell to string value
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "      "; // Return an empty string if the cell is null
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((int) cell.getNumericCellValue());
        } else {
            return "    ";
        }
    }}

