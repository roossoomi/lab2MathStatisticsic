package org.example;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class Main {
    private static final Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) {
        JTextField textField;
        JTextArea textArea;
        JComboBox<String> sheetComboBox;
        JFrame frame = new JFrame("Математическая статистика");
        frame.setSize(600, 400);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(null);

        /*textField = new JTextField();
        textField.setBounds(100, 100, 200, 30);
        frame.add(textField);*/

        JButton importButton = new JButton("Загрузить exel файл");
        importButton.setBounds(100, 30, 200, 30);
        frame.add(importButton);

        JButton exportButton = new JButton("Создать файл с результатами");
        exportButton.setBounds(100, 65, 200, 30);
        frame.add(exportButton);

        JButton showDataButton = new JButton("Показать данные");
        showDataButton.setBounds(100, 135, 200, 30);
        frame.add(showDataButton);

        sheetComboBox = new JComboBox<>();
        sheetComboBox.setBounds(320, 100, 200, 30);
        frame.add(sheetComboBox);

        textArea = new JTextArea();
        textArea.setLineWrap(true);
        textArea.setWrapStyleWord(true);
        textArea.setEditable(false);

        JScrollPane scrollPane = new JScrollPane(textArea);
        scrollPane.setBounds(50, 170, 500, 150);
        frame.add(scrollPane);

        frame.setVisible(true);

        importButton.addActionListener(new ActionListener() {
            HashMap<String, HashMap<String, Double>> mapArray = new HashMap<>();
            HashMap<String, Double> mapCovariance = new HashMap<>();

            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
                int result = fileChooser.showOpenDialog(null);

                if (result == JFileChooser.APPROVE_OPTION) {
                    File file = fileChooser.getSelectedFile();
                    try {
                        FileInputStream fis = new FileInputStream(file);
                        XSSFWorkbook workbook = new XSSFWorkbook(fis);
                        sheetComboBox.removeAllItems(); // Очистить выпадающий список

                        // Заполнить выпадающий список именами листов
                        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                            sheetComboBox.addItem(workbook.getSheetName(i));
                        }

                        sheetComboBox.addActionListener(new ActionListener() {
                            @Override
                            public void actionPerformed(ActionEvent e) {
                                String selectedSheetName = (String) sheetComboBox.getSelectedItem();
                                if (selectedSheetName != null) {
                                    XSSFSheet selectedSheet = workbook.getSheet(selectedSheetName);
                                    HashMap<String, double[]> commonMap = InputModule.extractDataFromSheet(selectedSheet);

                                    for (String key : commonMap.keySet()) {
                                        mapArray.put(key, new HashMap<>());
                                    }
                                    for (String key : commonMap.keySet()) {
                                        Calculation calculationExample = new Calculation(mapArray.get(key), commonMap.get(key));
                                        mapArray.put(key, calculationExample.makeCalculation(mapArray.get(key), commonMap.get(key)));
                                    }
                                    mapCovariance = Covarianc.getCovariance(commonMap);

                                    exportButton.addActionListener(new ActionListener() {
                                        @Override
                                        public void actionPerformed(ActionEvent e) {
                                            ExportModule.writeExcelFile(mapArray, mapCovariance);
                                            frame.dispose();
                                        }
                                    });

                                    showDataButton.addActionListener(new ActionListener() {
                                        @Override
                                        public void actionPerformed(ActionEvent e) {
                                            StringBuilder data = new StringBuilder();
                                            for (String key : mapArray.keySet()) {
                                                data.append("Случайная величина: ").append(key).append("\n");
                                                for (Map.Entry<String, Double> entry : mapArray.get(key).entrySet()) {
                                                    data.append(entry.getKey()).append(": ").append(entry.getValue()).append("\n");
                                                }
                                                data.append("\n");
                                            }
                                            data.append("Ковариация:\n");
                                            for (String key : mapCovariance.keySet()) {
                                                data.append(key).append(": ").append(mapCovariance.get(key)).append("\n");
                                            }
                                            textArea.setText(data.toString());
                                        }
                                    });
                                }
                            }
                        });

                    } catch (Exception ex) {
                        ex.printStackTrace();
                        logger.error("Ошибка при импорте и чтении данных из файла Excel", ex.getMessage());
                        String errorMessage = "Ошибка при импорте и чтении данных из файла Excel";
                        JOptionPane.showMessageDialog(null, errorMessage, "Ошибка", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
            }
        });
    }
}