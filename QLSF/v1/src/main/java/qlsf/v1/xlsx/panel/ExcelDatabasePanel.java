package qlsf.v1.xlsx.panel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

@Component
public class ExcelDatabasePanel extends JPanel {

    private final JButton openButton;
    private final JButton addReportButton;
    private final JButton saveButton;
    private final JButton saveAndExitButton;
    private final JButton changeSheetButton;
    private final JButton addSheetButton;
    private final JTable table;
    private final DefaultTableModel tableModel;
    private File currentFile;
    private Workbook workbook;
    private Sheet currentSheet;

    public ExcelDatabasePanel() {
        setLayout(new BorderLayout());

        openButton = new JButton("Open Excel File");
        addReportButton = new JButton("Add Report");
        saveButton = new JButton("Save");
        saveAndExitButton = new JButton("Save and Exit");
        changeSheetButton = new JButton("Change Sheet");
        addSheetButton = new JButton("Add Sheet");
        tableModel = new DefaultTableModel();
        table = new JTable(tableModel);

        openButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            int result = fileChooser.showOpenDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                if (selectedFile.getName().endsWith(".xlsx")) {
                    try {
                        currentFile = selectedFile;
                        readExcelFile(selectedFile);
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(null, "Error reading the Excel file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Please select a valid .xlsx file", "Invalid File", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        addReportButton.addActionListener(e -> addReport());

        saveButton.addActionListener(e -> {
            try {
                saveExcelFile();
                generatePaymentSummary();
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Error saving the Excel file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        });

        saveAndExitButton.addActionListener(e -> {
            try {
                saveExcelFile();
                generatePaymentSummary();
                System.exit(0);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, "Error saving the Excel file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        });

        changeSheetButton.addActionListener(e -> changeSheet());

        addSheetButton.addActionListener(e -> addSheet());

        JPanel buttonPanel = new JPanel();
        buttonPanel.add(openButton);
        buttonPanel.add(addReportButton);
        buttonPanel.add(saveButton);
        buttonPanel.add(saveAndExitButton);
        buttonPanel.add(changeSheetButton);
        buttonPanel.add(addSheetButton);

        add(buttonPanel, BorderLayout.NORTH);
        add(new JScrollPane(table), BorderLayout.CENTER);
    }

    private void readExcelFile(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            workbook = new XSSFWorkbook(fis);
            currentSheet = workbook.getSheetAt(0);
            loadSheetData(currentSheet);
        }
    }

    private void loadSheetData(Sheet sheet) {
        tableModel.setRowCount(0);
        tableModel.setColumnCount(0);

        Iterator<Row> rowIterator = sheet.iterator();
        if (rowIterator.hasNext()) {
            Row headerRow = rowIterator.next();
            Iterator<Cell> cellIterator = headerRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                tableModel.addColumn(cell.getStringCellValue());
            }
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            Object[] rowData = new Object[row.getLastCellNum()];
            int cellIndex = 0;
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case STRING:
                        rowData[cellIndex++] = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        rowData[cellIndex++] = cell.getNumericCellValue();
                        break;
                    default:
                        rowData[cellIndex++] = "";
                        break;
                }
            }
            tableModel.addRow(rowData);
        }
    }

    private void addReport() {
        String[] requiredColumns = {"Conductor", "Camión", "Fecha", "Pago", "Ubicación de salida", "Ubicación de llegada", "Seguro", "Libro digital", "Especificaciones", "Especificaciones Valor", "Especificación Fecha"};
        for (String column : requiredColumns) {
            if (tableModel.findColumn(column) == -1) {
                tableModel.addColumn(column);
            }
        }
    
        String conductor = JOptionPane.showInputDialog("Ingrese el nombre del Conductor:");
        String camion = JOptionPane.showInputDialog("Ingrese el nombre del Camión:");
        String fecha = JOptionPane.showInputDialog("Ingrese la Fecha:");
        String pago = JOptionPane.showInputDialog("Ingrese el Pago:");
        String ubicacionSalida = JOptionPane.showInputDialog("Ingrese la Ubicación de salida:");
        String ubicacionLlegada = JOptionPane.showInputDialog("Ingrese la Ubicación de llegada:");
        String seguro = JOptionPane.showInputDialog("Ingrese el Seguro:");
        String libroDigital = JOptionPane.showInputDialog("Ingrese el Libro digital:");
        String especificaciones = JOptionPane.showInputDialog("Ingrese las Especificaciones:");
        String especificacionesValor = JOptionPane.showInputDialog("Ingrese el valor de las Especificaciones:");
        String especificacionFecha = JOptionPane.showInputDialog("Ingrese la Especificación Fecha:");
    
        Object[] rowData = new Object[tableModel.getColumnCount()];
        for (int i = 0; i < rowData.length; i++) {
            rowData[i] = "";
        }
    
        rowData[tableModel.findColumn("Conductor")] = conductor;
        rowData[tableModel.findColumn("Camión")] = camion;
        rowData[tableModel.findColumn("Fecha")] = fecha;
        rowData[tableModel.findColumn("Pago")] = pago;
        rowData[tableModel.findColumn("Ubicación de salida")] = ubicacionSalida;
        rowData[tableModel.findColumn("Ubicación de llegada")] = ubicacionLlegada;
        rowData[tableModel.findColumn("Seguro")] = seguro;
        rowData[tableModel.findColumn("Libro digital")] = libroDigital;
        rowData[tableModel.findColumn("Especificaciones")] = especificaciones;
        rowData[tableModel.findColumn("Especificaciones Valor")] = especificacionesValor;
        rowData[tableModel.findColumn("Especificación Fecha")] = especificacionFecha;
    
        tableModel.addRow(rowData);
    }

    private void saveExcelFile() throws IOException {
        if (currentFile == null) {
            JOptionPane.showMessageDialog(null, "No file is currently open.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        try (FileOutputStream fos = new FileOutputStream(currentFile)) {
            // Crear una nueva hoja con el mismo nombre que la hoja actual
            String sheetName = currentSheet.getSheetName();
            int sheetIndex = workbook.getSheetIndex(currentSheet);
            if (sheetIndex != -1) {
                workbook.removeSheetAt(sheetIndex);
            }
            currentSheet = workbook.createSheet(sheetName);

            // Escribir los datos en la nueva hoja
            Row headerRow = currentSheet.createRow(0);
            for (int col = 0; col < tableModel.getColumnCount(); col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(tableModel.getColumnName(col));
            }

            for (int row = 0; row < tableModel.getRowCount(); row++) {
                Row excelRow = currentSheet.createRow(row + 1);
                for (int col = 0; col < tableModel.getColumnCount(); col++) {
                    Cell cell = excelRow.createCell(col);
                    Object value = tableModel.getValueAt(row, col);
                    if (value != null) {
                        if (value instanceof String) {
                            cell.setCellValue((String) value);
                        } else if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else {
                            cell.setCellValue(value.toString());
                        }
                    } else {
                        cell.setCellValue("");
                    }
                }
            }

            workbook.write(fos);
        }
    }

    private void changeSheet() {
        String[] sheetNames = new String[workbook.getNumberOfSheets()];
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheetNames[i] = workbook.getSheetName(i);
        }

        String selectedSheet = (String) JOptionPane.showInputDialog(null, "Select Sheet", "Change Sheet",
                JOptionPane.QUESTION_MESSAGE, null, sheetNames, sheetNames[0]);

        if (selectedSheet != null) {
            currentSheet = workbook.getSheet(selectedSheet);
            loadSheetData(currentSheet);
        }
    }

    private void addSheet() {
        String sheetName = JOptionPane.showInputDialog("Enter the name of the new sheet:");
        if (sheetName != null && !sheetName.trim().isEmpty()) {
            currentSheet = workbook.createSheet(sheetName);
            loadSheetData(currentSheet);
        }
    }

    private void generatePaymentSummary() throws IOException {
        Map<String, Double> paymentSummary = new HashMap<>();
        Map<String, Double> twelvePercentSummary = new HashMap<>();
        Map<String, Double> comparisonSummary = new HashMap<>();
        int conductorColumn = tableModel.findColumn("Conductor");
        int pagoColumn = tableModel.findColumn("Pago");
        int seguroColumn = tableModel.findColumn("Seguro");
        int libroDigitalColumn = tableModel.findColumn("Libro digital");
        int especificacionesValorColumn = tableModel.findColumn("Especificaciones Valor");

        if (conductorColumn == -1 || pagoColumn == -1 || seguroColumn == -1 || libroDigitalColumn == -1 || especificacionesValorColumn == -1) {
            JOptionPane.showMessageDialog(null, "Columns 'Conductor', 'Pago', 'Seguro', 'Libro digital', or 'Especificaciones Valor' not found.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Solicitar el porcentaje al usuario
        String percentageStr = JOptionPane.showInputDialog("Ingrese el porcentaje que desea calcular de Pago:");
        double percentage;
        try {
            percentage = Double.parseDouble(percentageStr);
        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(null, "Porcentaje inválido. Por favor, ingrese un número válido.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        for (int row = 0; row < tableModel.getRowCount(); row++) {
            String conductor = (String) tableModel.getValueAt(row, conductorColumn);
            Object pagoValue = tableModel.getValueAt(row, pagoColumn);
            Object seguroValue = tableModel.getValueAt(row, seguroColumn);
            Object libroDigitalValue = tableModel.getValueAt(row, libroDigitalColumn);
            Object especificacionesValorValue = tableModel.getValueAt(row, especificacionesValorColumn);

            Double pago = 0.0;
            Double seguro = 0.0;
            Double libroDigital = 0.0;
            Double especificacionesValor = 0.0;

            if (pagoValue instanceof Number) {
                pago = ((Number) pagoValue).doubleValue();
            } else if (pagoValue instanceof String) {
                try {
                    pago = Double.parseDouble((String) pagoValue);
                } catch (NumberFormatException e) {
                    JOptionPane.showMessageDialog(null, "Invalid payment value at row " + (row + 1), "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }
            }

            if (seguroValue instanceof Number) {
                seguro = ((Number) seguroValue).doubleValue();
            } else if (seguroValue instanceof String) {
                try {
                    seguro = Double.parseDouble((String) seguroValue);
                } catch (NumberFormatException e) {
                    JOptionPane.showMessageDialog(null, "Invalid seguro value at row " + (row + 1), "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }
            }

            if (libroDigitalValue instanceof Number) {
                libroDigital = ((Number) libroDigitalValue).doubleValue();
            } else if (libroDigitalValue instanceof String) {
                try {
                    libroDigital = Double.parseDouble((String) libroDigitalValue);
                } catch (NumberFormatException e) {
                    JOptionPane.showMessageDialog(null, "Invalid libro digital value at row " + (row + 1), "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }
            }

            if (especificacionesValorValue instanceof Number) {
                especificacionesValor = ((Number) especificacionesValorValue).doubleValue();
            } else if (especificacionesValorValue instanceof String) {
                try {
                    especificacionesValor = Double.parseDouble((String) especificacionesValorValue);
                } catch (NumberFormatException e) {
                    JOptionPane.showMessageDialog(null, "Invalid especificaciones valor value at row " + (row + 1), "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }
            }

            Double totalPago = (pago * 0.88) - seguro - libroDigital - especificacionesValor;
            paymentSummary.put(conductor, paymentSummary.getOrDefault(conductor, 0.0) + totalPago);
            twelvePercentSummary.put(conductor, twelvePercentSummary.getOrDefault(conductor, 0.0) + (pago * 0.12));
            comparisonSummary.put(conductor, comparisonSummary.getOrDefault(conductor, 0.0) + (pago * (percentage / 100)));
        }

        int summaryStartColumn = tableModel.findColumn("Conductor Pago");
        if (summaryStartColumn == -1) {
            summaryStartColumn = tableModel.getColumnCount();
            tableModel.addColumn("Conductor Pago");
            tableModel.addColumn("Total Pago");
            tableModel.addColumn("12%");
            tableModel.addColumn("Comparacion");
        }

        for (int row = 0; row < tableModel.getRowCount(); row++) {
            tableModel.setValueAt("", row, summaryStartColumn);
            tableModel.setValueAt("", row, summaryStartColumn + 1);
            tableModel.setValueAt("", row, summaryStartColumn + 2);
            tableModel.setValueAt("", row, summaryStartColumn + 3);
        }

        int rowIndex = 0;
        for (Map.Entry<String, Double> entry : paymentSummary.entrySet()) {
            tableModel.setValueAt(entry.getKey(), rowIndex, summaryStartColumn);
            tableModel.setValueAt(entry.getValue(), rowIndex, summaryStartColumn + 1);
            tableModel.setValueAt(twelvePercentSummary.get(entry.getKey()), rowIndex, summaryStartColumn + 2);
            tableModel.setValueAt(comparisonSummary.get(entry.getKey()), rowIndex, summaryStartColumn + 3);
            rowIndex++;
        }

        try (FileOutputStream fos = new FileOutputStream(currentFile)) {
            workbook.write(fos);
        }
    }
}