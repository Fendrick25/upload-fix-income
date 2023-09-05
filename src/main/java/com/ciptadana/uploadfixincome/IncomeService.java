package com.ciptadana.uploadfixincome;

import jakarta.annotation.PostConstruct;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class IncomeService {
    String excelFilePath = "Upload_Fixed_Income_20230906.xls";
    String imageOutputPath = "excel/output.jpeg";
    BufferedImage IDX;
    BufferedImage KPEI;
    BufferedImage KSEI;

    public void init() throws IOException{
        IDX = ImageIO.read(new File("asset/IDX.png"));
        KPEI = ImageIO.read(new File("asset/KPEI.jpg"));
        KSEI = ImageIO.read(new File("asset/KSEI.png"));
    }

   // @PostConstruct
    public void convertExcelToImage2() throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(excelFilePath))) {

            Workbook workbook;
            if (excelFilePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else if (excelFilePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("Invalid file type");
            }

            Sheet sheet = workbook.getSheetAt(0);

            int totalWidth = 0;
            short maxColumns = sheet.getRow(0).getLastCellNum();
            for (int i = 0; i < maxColumns; i++) {
                totalWidth += sheet.getColumnWidthInPixels(i);
            }

            int totalHeight = 0;
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                totalHeight += sheet.getRow(i).getHeight();
            }

            BufferedImage image = new BufferedImage(totalWidth, totalHeight, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = image.createGraphics();
            graphics.setColor(Color.WHITE);
            graphics.fillRect(0, 0, totalWidth, totalHeight);

            graphics.setFont(new Font("Arial", Font.PLAIN, 10)); // Adjust font size and style as needed

            int yPosition = 0;
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                int xPosition = 0;
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        int cellWidth = (int) sheet.getColumnWidthInPixels(j);
                        int cellHeight = row.getHeight();
                        String cellValue = new DataFormatter().formatCellValue(cell);
                        System.out.println(cellValue);
                        graphics.setColor(Color.BLACK);
                        graphics.drawRect(xPosition, yPosition, cellWidth, cellHeight);
                        graphics.drawString(cellValue, xPosition + 5, yPosition + cellHeight - 5); // Add some padding
                    }
                    xPosition += sheet.getColumnWidthInPixels(j);
                }
                yPosition += sheet.getRow(i).getHeight();
            }

            graphics.dispose();
            ImageIO.write(image, "jpeg", new File(imageOutputPath));
        }
    }

    @PostConstruct
    public void convertExcelToImage3() throws IOException {
        init();
        try (FileInputStream fis = new FileInputStream(new File(excelFilePath))) {

            Workbook workbook;
            if (excelFilePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else if (excelFilePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("Invalid file type");
            }

            Sheet sheet = workbook.getSheetAt(0);

            int totalWidth = 0;
            short maxColumns = sheet.getRow(0).getLastCellNum();
            for (int i = 0; i < maxColumns; i++) {
                totalWidth += sheet.getColumnWidthInPixels(i);
            }

            int totalHeight = 0;
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                totalHeight += sheet.getRow(i).getHeight();
            }

            BufferedImage image = new BufferedImage(totalWidth, totalHeight, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = image.createGraphics();
            graphics.setColor(Color.WHITE);
            graphics.fillRect(0, 0, totalWidth, totalHeight);

            graphics.setFont(new Font("Arial", Font.PLAIN, 10)); // Adjust font size and style as needed

            int yPosition = 0;
            List<FixIncome> fixIncomes = new ArrayList<>();
            for (int i = 3; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                int xPosition = 0;
                FixIncome fixIncome = FixIncome.builder()
                        .issuerName(new DataFormatter().formatCellValue(row.getCell(0)))
                        .category(new DataFormatter().formatCellValue(row.getCell(1)))
                        .type(new DataFormatter().formatCellValue(row.getCell(2)))
                        .coupon(new DataFormatter().formatCellValue(row.getCell(3)))
                        .rating(new DataFormatter().formatCellValue(row.getCell(4)))
                        .maturity(new DataFormatter().formatCellValue(row.getCell(5)))
                        .bidPrice(new DataFormatter().formatCellValue(row.getCell(6)))
                        .offerPrice(new DataFormatter().formatCellValue(row.getCell(7)))
                        .yieldBid(new DataFormatter().formatCellValue(row.getCell(8)))
                        .yieldOffer(new DataFormatter().formatCellValue(row.getCell(9)))
                        .currency(new DataFormatter().formatCellValue(row.getCell(10)))
                        .accountMin(new DataFormatter().formatCellValue(row.getCell(11)))
                        .build();

                fixIncomes.add(fixIncome);

/*                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        int cellWidth = (int) sheet.getColumnWidthInPixels(j);
                        int cellHeight = row.getHeight();
                        String cellValue = new DataFormatter().formatCellValue(cell);


                        if(!cellValue.equalsIgnoreCase("")){
                            //System.out.println(cellValue);
                            graphics.setColor(Color.BLACK);
                            graphics.drawRect(xPosition, yPosition, cellWidth, cellHeight);
                            graphics.drawString(cellValue, xPosition + 5, yPosition + cellHeight - 5); // Add
                        }
                    }
                    xPosition += sheet.getColumnWidthInPixels(j);
                }
                yPosition += 30;*/
            }

            int cellHeight = 30;
            int textPadding = 5;
            int headerHeight = 30; // or any value suitable for your needs
            int headerPadding = 5; // p

            HashMap<String, HashMap<String, List<FixIncome>>> fixIncomeByCurrencyAndType =
                    fixIncomes.stream()
                            .collect(Collectors.groupingBy(
                                    FixIncome::getCurrency,
                                    HashMap::new,
                                    Collectors.groupingBy(FixIncome::getType, HashMap::new, Collectors.toList())
                            ));

            int cellColumn = 0;
            for (String currency : fixIncomeByCurrencyAndType.keySet()) {

                Row headerRow = sheet.getRow(2);
                //String[] headers = new String[headerRow.getLastCellNum()];
                List<String> headers = new ArrayList<>();
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                  if(!(i == 1 || i == 2 | i == 4)){
                      headers.add(new DataFormatter().formatCellValue(headerRow.getCell(i)));
                  }
                }


                // 2. Draw the headers

                HashMap<String, List<FixIncome>> fixIncomeMap = fixIncomeByCurrencyAndType.get(currency);
                for(String type : fixIncomeMap.keySet()){
                    List<FixIncome> incomes = fixIncomeMap.get(type);
                    int xPositionHeader = 0;

                    for (int j = 0; j < headers.size(); j++) {
                        cellColumn += (int) sheet.getColumnWidthInPixels(j);
                    }
                    yPosition += cellHeight;
                    for (int j = 0; j < headers.size(); j++) {
                        int cellWidth = (int) sheet.getColumnWidthInPixels(j);
                        if(j == 0){
                            graphics.setColor(Color.BLACK);
                            graphics.drawRect(xPositionHeader, yPosition, cellColumn, cellHeight);

                            graphics.setColor(Color.BLACK);
                            System.out.println( type);
                            graphics.drawString(generateTitle(type).concat(" " + currency.toUpperCase()), xPositionHeader + textPadding, yPosition + cellHeight - textPadding);
                            yPosition += cellHeight;
                        }

                        graphics.setColor(Color.BLACK);
                        graphics.drawRect(xPositionHeader, yPosition, cellWidth, cellHeight);

                        graphics.setColor(Color.BLACK);
                        graphics.drawString(headers.get(j), xPositionHeader + textPadding, yPosition + cellHeight - textPadding);

                        xPositionHeader += cellWidth;
                    }
                    yPosition += cellHeight;



                    for (FixIncome income : incomes) {
                        int xPosition = 0;
                        for (int j = 0; j < 12; j++) { // 12 fields in FixIncome

                            if(!(j == 1 || j == 2 | j == 4)){
                                int cellWidth = (int) sheet.getColumnWidthInPixels(j);
                                graphics.setColor(Color.BLACK);
                                graphics.drawRect(xPosition, yPosition, cellWidth, cellHeight);

                                graphics.setColor(Color.BLACK);
                                String data = getFieldByIndex(income, j);
                                //System.out.println(data);
                                graphics.drawString(data, xPosition + textPadding, yPosition + cellHeight - textPadding);

                                xPosition += cellWidth;
                            }
                        }
                        yPosition += cellHeight;
                        graphics.drawString(" ", xPosition + textPadding, yPosition + cellHeight - textPadding);
                    }
                   // yPosition += cellHeight;
                }
            }

/*            // Draw image at the top-right corner
            graphics.drawImage(IDX, canvasWidth, 0, null);

            // Draw image at the bottom-left corner
            graphics.drawImage(IDX, 0, canvasHeight, null);

            // Draw image at the bottom-right corner*/



            cellColumn = cellColumn / 3;
            graphics.drawImage(IDX,  300, cellColumn - 100, null);
            graphics.drawImage(KPEI, 500, cellColumn - 100, null);
            graphics.drawImage(KSEI, 700, cellColumn - 100, null);


            graphics.dispose();
            ImageIO.write(image, "jpeg", new File(imageOutputPath));


        }


    }

    private String getFieldByIndex(FixIncome income, int index) {
        switch(index) {
            case 0: return income.getIssuerName();
            case 1: return income.getCategory();
            case 2: return income.getType();
            case 3: return income.getCoupon();
            case 4: return income.getRating();
            case 5: return income.getMaturity();
            case 6: return income.getBidPrice();
            case 7: return income.getOfferPrice();
            case 8: return income.getYieldBid();
            case 9: return income.getYieldOffer();
            case 10: return income.getCurrency();
            case 11: return income.getAccountMin();
            default: return "";
        }
    }

    private String generateTitle(String type){
        if (type.toLowerCase().contains("quarterly")){
            return "LOCAL BOND (QUARTERLY)";
        }

        if (type.toLowerCase().contains("semi annually")){
            return "LOCAL BOND (SEMI ANNUALLY)";
        }

        return "UNKNOWN";
    }
}
