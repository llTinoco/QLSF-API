package qlsf.v1;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import qlsf.v1.xlsx.panel.ExcelDatabasePanel;

import javax.swing.*;
import java.awt.*;

@SpringBootApplication
public class V1Application {

    public static void main(String[] args) {
        // Configura el entorno gráfico
        if (!GraphicsEnvironment.isHeadless()) {
            System.setProperty("java.awt.headless", "false");
        }
        SpringApplication.run(V1Application.class, args);

        // Mostrar el panel en una ventana
        EventQueue.invokeLater(() -> {
            JFrame frame = new JFrame("Excel Database Panel");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.getContentPane().add(new ExcelDatabasePanel());
            frame.pack();
            frame.setLocationRelativeTo(null);
            frame.setVisible(true);
        });
    }

    @Bean
    public ExcelDatabasePanel excelDatabasePanel() {
        return new ExcelDatabasePanel();
    }
}