package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException, SQLException {
        String filePath = "C:\\Users\\lukin\\OneDrive\\Área de Trabalho\\Nova pasta\\planilha.xlsx";
        String tableName = "product";
        String url = "jdbc:mysql://localhost:3306/db000";
        String username = "root";
        String password = "@soma+";
        String defaultValue = "";

        try (Connection connection = DriverManager.getConnection(url, username, password)) {
            FileInputStream fileInputStream = new FileInputStream(new File(filePath));
            Workbook workbook = WorkbookFactory.create(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter dataFormatter = new DataFormatter();

            StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (barcode, name, cost, price, current_stock");
            StringBuilder valuePlaceholders = new StringBuilder(" VALUES (?, ?, ?, ?, ?");
            List<String> defaultValues = new ArrayList<>();

            DatabaseMetaData metaData = connection.getMetaData();
            ResultSet resultSet = metaData.getColumns(null, null, tableName, null);

            int totalColumnsInDatabase = 5;
/*            ResultSetMetaData rsmd = resultSet.getMetaData();
            System.out.println("querying SELECT * FROM XXX");
            int columnsNumber = rsmd.getColumnCount();
            while (resultSet.next()) {
                for (int i = 1; i <= columnsNumber; i++) {
                    if (i > 1) System.out.print(",  ");
                    String columnValue = resultSet.getString(i);
                    System.out.print(columnValue + " " + rsmd.getColumnName(i));
                }
                System.out.println("");
            }*/
            while (resultSet.next()) {
                String columnName = resultSet.getString("COLUMN_NAME");
                if (!columnName.equals("barcode") && !columnName.equals("name") && !columnName.equals("cost") && !columnName.equals("price") && !columnName.equals("current_stock")) {
                    if (!columnName.equals("id")) {
                        if (totalColumnsInDatabase > 0) {
                            insertQuery.append(", ");
                            valuePlaceholders.append(", ");
                    }
                    insertQuery.append(", ").append(columnName);
                    valuePlaceholders.append(", ?");
                    defaultValues.add(defaultValue);
                    totalColumnsInDatabase++;
                    System.out.println(columnName);
                }
            }
            resultSet.close();
            insertQuery.append(")");
            valuePlaceholders.append(")");
            insertQuery.append(valuePlaceholders);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell barcodeCell = row.getCell(0);
                Cell nameCell = row.getCell(1);
                Cell costCell = row.getCell(2);
                Cell priceCell = row.getCell(3);
                Cell currentStockCell = row.getCell(4);

                if (barcodeCell != null && nameCell != null && costCell != null && priceCell != null && currentStockCell != null) {
                    String barcodeValue = dataFormatter.formatCellValue(barcodeCell);
                    String nameValue = dataFormatter.formatCellValue(nameCell);
                    double costValue = costCell.getNumericCellValue();
                    double priceValue = priceCell.getNumericCellValue();
                    double currentStockValue = currentStockCell.getNumericCellValue();

                    PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString());
                    preparedStatement.setString(1, barcodeValue);
                    preparedStatement.setString(2, nameValue);
                    preparedStatement.setDouble(3, costValue);
                    preparedStatement.setDouble(4, priceValue);
                    preparedStatement.setDouble(5, currentStockValue);

                    // Preencha os valores padrão para as colunas NOT NULL adicionais
                    for (int j = 0; j < defaultValues.size(); j++) {
                        preparedStatement.setString(j + 5, defaultValues.get(j));
                    }
                    preparedStatement.executeUpdate();
                    preparedStatement.close();
                } else {
                    // Trate o caso em que alguma das células necessárias está vazia
                }
            }

            connection.close();
            System.out.println("Dados inseridos");
        } }catch (Exception e) {
            e.printStackTrace();
        }
    }
}