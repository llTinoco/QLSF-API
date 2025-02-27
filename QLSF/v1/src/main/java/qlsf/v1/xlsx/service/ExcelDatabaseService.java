package qlsf.v1.xlsx.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelDatabaseService {

    private static final Logger logger = LoggerFactory.getLogger(ExcelDatabaseService.class);

    public List<String[]> leerDatos(String filePath) throws IOException {
        List<String[]> datos = new ArrayList<>();
        logger.info("Leyendo datos desde el archivo: {}", filePath);

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

        logger.info("Datos leídos correctamente desde el archivo.");
        return datos;
    }

    public void escribirDatos(String filePath, List<String[]> datos) throws IOException {
        logger.info("Escribiendo datos en el archivo: {}", filePath);

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

        logger.info("Datos escritos correctamente en el archivo.");
    }
}