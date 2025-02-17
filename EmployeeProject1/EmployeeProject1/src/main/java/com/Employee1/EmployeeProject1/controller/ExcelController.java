package com.Employee1.EmployeeProject1.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.util.*;

public class ExcelController {

    private static final String INPUT_FILE = "InputSheet.xlsx"; // Input file in resources
    private static final String OUTPUT_FILE = "target/classes/copied.xlsx"; // Output file location

    // Fixed column names where we will search for the skill
    private static final List<String> SEARCH_COLUMNS = Arrays.asList("skill1", "skill2", "skill3");

    public String filterAndSaveExcel(String skill) {
        try {
            // Load InputSheet.xlsx from resources folder
            File file = new ClassPathResource(INPUT_FILE).getFile();
            FileInputStream inputStream = new FileInputStream(file);
            Workbook inputWorkbook = new XSSFWorkbook(inputStream);

            // Create output workbook
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Filtered Data");

            // Ensure the output directory exists
            File outputFile = new File(OUTPUT_FILE);
            File parentDir = outputFile.getParentFile();
            if (parentDir != null && !parentDir.exists()) {
                parentDir.mkdirs(); // Create the directory if it doesn't exist
            }

            boolean headerCopied = false; // Ensure header is copied once

            for (int i = 0; i < inputWorkbook.getNumberOfSheets(); i++) {
                Sheet inputSheet = inputWorkbook.getSheetAt(i);

                // Find indexes of the specified columns
                List<Integer> columnIndexes = findColumnIndexes(inputSheet, SEARCH_COLUMNS);
                if (columnIndexes.isEmpty()) {
                    System.out.println("No matching columns found in the Excel file!");
                    return "No matching columns found!";
                }

//                System.out.println("🔍 Found columns at indexes: " + columnIndexes);

                // Copy matching rows
                for (Row inputRow : inputSheet) {
                    if (inputRow.getRowNum() == 0) {
                        if (!headerCopied) {
                            copyRow(inputRow, outputSheet.createRow(0)); // Copy header once
                            headerCopied = true;
                        }
                        continue;
                    }

                    // Check if any specified column contains the exact skill
                    for (int colIndex : columnIndexes) {
                        Cell cell = inputRow.getCell(colIndex);
                        if (cell != null) {
                            String cellValue = getCellValueAsString(cell);
//                            System.out.println("🔎 Checking row " + inputRow.getRowNum() + " in column " + colIndex + ": " + cellValue);

                            if (isExactMatch(cellValue, skill)) {
//                                System.out.println("Exact match found! Copying row " + inputRow.getRowNum());
                                copyRow(inputRow, outputSheet.createRow(outputSheet.getPhysicalNumberOfRows()));
                                break; // Copy row only once if match is found
                            }
                        }
                    }
                }
            }

            // Save the filtered file
            try (FileOutputStream fileOut = new FileOutputStream(OUTPUT_FILE)) {
                outputWorkbook.write(fileOut);
            }

            return "Filtered data saved at: " + OUTPUT_FILE;

        } catch (IOException e) {
            e.printStackTrace();
            return "Error processing the Excel file.";
        }
    }

    // Find indexes of specified columns
    private List<Integer> findColumnIndexes(Sheet sheet, List<String> columnNames) {
        List<Integer> indexes = new ArrayList<>();
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return indexes;

        for (Cell cell : headerRow) {
            if (cell.getCellType() == CellType.STRING) {
                String colName = cell.getStringCellValue().trim().toLowerCase();
//                System.out.println("📝 Found column: " + colName);

                if (columnNames.contains(colName)) {
                    indexes.add(cell.getColumnIndex());
                }
            }
        }
        return indexes;
    }

    // Copy an entire row
    private void copyRow(Row inputRow, Row outputRow) {
        for (Cell inputCell : inputRow) {
            Cell outputCell = outputRow.createCell(inputCell.getColumnIndex(), inputCell.getCellType());

            switch (inputCell.getCellType()) {
                case STRING:
                    outputCell.setCellValue(inputCell.getStringCellValue());
                    break;
                case NUMERIC:
                    outputCell.setCellValue(inputCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    outputCell.setCellValue(inputCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    outputCell.setCellFormula(inputCell.getCellFormula());
                    break;
                case BLANK:
                    outputCell.setBlank();
                    break;
                default:
                    break;
            }
        }
    }

    // Get cell value as string (handles different types)
    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    // Check if cell value matches the skill exactly
    private boolean isExactMatch(String cellValue, String skill) {
        if (cellValue == null || cellValue.trim().isEmpty()) return false;

        // Split multiple skills by commas, spaces, or slashes
        String[] skills = cellValue.toLowerCase().split("[, /]+");

        for (String s : skills) {
            if (s.trim().equals(skill.toLowerCase().trim())) {
                return true; // Exact match found
            }
        }
        return false;
    }
}
