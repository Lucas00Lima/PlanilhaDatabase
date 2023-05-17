package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class Main {
    public static void main(String[] args) throws IOException, SQLException {
        String filePath = "caminho\\arquivo.xlsx";
        String tableName = "product";
        String url = "jdbc:mysql://localhost:3306/banco";
        String username = "root";
        String password = "";
        String defaultValue = "";
        try (Connection connection = DriverManager.getConnection(url, username, password)) {
            connection.setAutoCommit(false);
            int batchSize = 9000;
            int count = 0;
            FileInputStream fileInputStream = new FileInputStream(new File(filePath));
            Workbook workbook = WorkbookFactory.create(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter dataFormatter = new DataFormatter();
            StringBuilder insertQuery = new StringBuilder("INSERT INTO " + tableName + " (barcode, name, cost, price, current_stock");
            StringBuilder valuePlaceholders = new StringBuilder(" VALUES (?, ?, ?, ?, ?");
            List<String> defaultValues = new ArrayList<>();
            DatabaseMetaData metaData = connection.getMetaData();
            ResultSet resultSet = metaData.getColumns(null, null, tableName, null);
            int totalColumnsInDatabase = 6;

            //Verificação e exclusão das colunas
            while (resultSet.next()) {
                String columnName = resultSet.getString("COLUMN_NAME");
                if (!columnName.equals("barcode") && !columnName.equals("name") && !columnName.equals("cost") && !columnName.equals("price") && !columnName.equals("current_stock")) {
                    if (!columnName.equals("id") && !columnName.equals("validity") && !columnName.equals("deleted_at") && !columnName.equals("delivery") && !columnName.equals("card") && !columnName.equals("balcony") && !columnName.equals("parameters")) {
                        if (totalColumnsInDatabase > 0) {
                            insertQuery.append(",");
                            valuePlaceholders.append(",");
                        }
                        insertQuery.append(columnName);
                        valuePlaceholders.append("?");
                        defaultValues.add(defaultValue);
                        totalColumnsInDatabase++;
                    }
                }
            }
            resultSet.close();
            insertQuery.append(")");
            valuePlaceholders.append(")");
            insertQuery.append(valuePlaceholders);

            //Separando as celulas da planilha.
            Set<String> nomesLidos = new HashSet<>();
            int i;
            for (i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell barcodeCell = row.getCell(0);
                Cell nameCell = row.getCell(1);
                Cell costCell = row.getCell(2);
                Cell priceCell = row.getCell(3);
                Cell currentStockCell = row.getCell(4);

                //Leitura de nome repetido
                String name = dataFormatter.formatCellValue(nameCell);
                if (nomesLidos.contains(name)) {
                    continue;
                }
                nomesLidos.add(name);
                //Leitura do estoque 0
                int currentStock = (int) currentStockCell.getNumericCellValue();
                if (currentStock == 0 || currentStock < 0) {
                    continue;
                }
                //Query
                if (barcodeCell != null && nameCell != null && costCell != null && priceCell != null && currentStockCell != null) {
                    String barcodeValue = dataFormatter.formatCellValue(barcodeCell);
                    String nameValue = dataFormatter.formatCellValue(nameCell);
                    int costValue = (int) (costCell.getNumericCellValue() * 100);
                    int priceValue = (int) (priceCell.getNumericCellValue() * 100);
                    int currentStockValue = (int) (currentStockCell.getNumericCellValue() * 1000);
                    PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString());
                    preparedStatement.setString(1, barcodeValue);
                    preparedStatement.setString(2, nameValue);
                    preparedStatement.setDouble(3, costValue);
                    preparedStatement.setDouble(4, priceValue);
                    preparedStatement.setDouble(5, currentStockValue);

                    //NOT NULL adicionais
                    for (int j = 0; j < defaultValues.size(); j++) {
                        String value = defaultValues.get(j);
                        if (value.isEmpty()) {
                            preparedStatement.setInt(j + 6, 0);
                        } else {
                            preparedStatement.setString(j + 6, value);
                        }
                    }
                    preparedStatement.executeUpdate();
                    count++;
                    if (count % batchSize == 0) {
                        connection.commit();
                    }
                    //Update do internal_code
                    String updateQuery = "UPDATE " + tableName + " SET internal_code = id";
                    preparedStatement.executeUpdate(updateQuery);
                    String updateDescription = "UPDATE " + tableName + " SET description = ''";
                    preparedStatement.executeUpdate(updateDescription);
                    preparedStatement.close();
                }
            }
            connection.commit();
            connection.close();
            System.out.println("Dados inseridos");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}