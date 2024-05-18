package org.example;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Main {
    private static final Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) {
        JTextField textField;
        JTextArea textArea;
        JComboBox<String> sheetComboBox;
        JFrame frame = new JFrame("Математическая статистика");
        frame.setSize(800, 380);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(null);


        JLabel mainInfo = new JLabel("<html>В данной программе вы можете рассчитать основные статистические значения.<br>Алгоритм работы:<br>1. Загрузить файл с данными<br>2. Выбрать нужный лист<br>3. Рассчитать данные или создать итоговый файл<br>Хорошей работы!</html>");
        mainInfo.setBounds(100, 10, 590, 100);
        frame.add(mainInfo);
        JButton importButton = new JButton("Загрузить excel файл");
        importButton.setBounds(100, 130, 220, 30);
        frame.add(importButton);

        JButton exportButton = new JButton("Создать файл с результатами");
        exportButton.setBounds(100, 180, 220, 30);
        frame.add(exportButton);

        JButton showDataButton = new JButton("Показать данные");
        showDataButton.setBounds(100, 230, 220, 30);
        frame.add(showDataButton);

        JButton showLog = new JButton("Показать журнал логирования");
        showLog.setBounds(100, 280, 220, 30);
        frame.add(showLog);

        sheetComboBox = new JComboBox<>();
        sheetComboBox.setBounds(350, 130, 200, 30);
        frame.add(sheetComboBox);

        textArea = new JTextArea();
        textArea.setLineWrap(true);
        textArea.setWrapStyleWord(true);
        textArea.setEditable(false);

        JScrollPane scrollPane = new JScrollPane(textArea);
        scrollPane.setBounds(350, 180, 300, 150);
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
                        sheetComboBox.removeAllItems();


                        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                            sheetComboBox.addItem(workbook.getSheetName(i));
                        }

                        sheetComboBox.addActionListener(new ActionListener() {
                            @Override
                            public void actionPerformed(ActionEvent e) {
                                String selectedSheetName = (String) sheetComboBox.getSelectedItem();
                                XSSFSheet selectedSheet = workbook.getSheet(selectedSheetName);

                                int rows = selectedSheet.getPhysicalNumberOfRows();
                                if (rows < 2) {
                                    logger.error("Попытка открыть лист без выборок");
                                    JOptionPane.showMessageDialog(null, "Лист пустой или содержит только заголовок", "Ошибка", JOptionPane.ERROR_MESSAGE);
                                }

                                Row headerRow = selectedSheet.getRow(0);
                                int cols = headerRow.getPhysicalNumberOfCells();

                                int[] columnRowCounts = new int[cols];

                                for (int i = 0; i < cols; i++) {
                                    for (int j = 1; j < rows; j++) {
                                        Row row = selectedSheet.getRow(j);
                                        if (row != null) {
                                            Cell cell = row.getCell(i);
                                            if (cell != null) {
                                                columnRowCounts[i]++;
                                            }
                                        }
                                    }
                                }

                                int rowCount = columnRowCounts[0];
                                for (int i = 1; i < cols; i++) {
                                    if (columnRowCounts[i] != rowCount) {
                                        logger.error("Попытка расчета с разным количеством элементов");
                                        JOptionPane.showMessageDialog(null, "В выборках разное количество элементов, дальнейшие расчеты невозможны", "Ошибка", JOptionPane.ERROR_MESSAGE);
                                        break;
                                    }
                                }

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
                                        //frame.dispose();
                                    }
                                });
                                showLog.addActionListener(new ActionListener() {
                                    @Override
                                    public void actionPerformed(ActionEvent e) {

                                        try {
                                            File logFile = new File(System.getProperty("user.dir") + "/LogMagazinee.txt");
                                            if (logFile.exists()) {
                                                Desktop.getDesktop().open(logFile);
                                            } else {
                                                JOptionPane.showMessageDialog(null, "Файл журнала не найден", "Ошибка", JOptionPane.ERROR_MESSAGE);
                                            }
                                        } catch (IOException ex) {
                                            JOptionPane.showMessageDialog(null, "Ошибка при открытии файла журнала", "Ошибка", JOptionPane.ERROR_MESSAGE);
                                        }
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