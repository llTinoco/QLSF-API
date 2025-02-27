package qlsf.v1.xlsx.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
public class ExcelDatabaseController {

    private String filePath;

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public List<String[]> leerDatos() throws IOException {
        if (filePath == null || filePath.isEmpty()) {
            throw new IllegalArgumentException("La ruta del archivo no está establecida.");
        }

        List<String[]> datos = new ArrayList<>();

        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                int numCells = row.getPhysicalNumberOfCells();
                String[] filaDatos = new String[numCells];
                for (int i = 0; i < numCells; i++) {
                    Cell cell = row.getCell(i);
                    filaDatos[i] = cell.toString();
                }
                datos.add(filaDatos);
            }
        }

        return datos;
    }

    public void escribirDatos(List<String[]> datos) throws IOException {
        if (filePath == null || filePath.isEmpty()) {
            throw new IllegalArgumentException("La ruta del archivo no está establecida.");
        }

        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            int rowNum = sheet.getLastRowNum() + 1;
            for (String[] filaDatos : datos) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 0; i < filaDatos.length; i++) {
                    row.createCell(i).setCellValue(filaDatos[i]);
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
        }
    }
}