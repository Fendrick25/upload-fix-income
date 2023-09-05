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
    BufferedImage idxImage;
    BufferedImage kpeiImage;
    BufferedImage kseiImage;
    BufferedImage ciptadanaImage;
    Color brown = new Color(154,102,51);
    Color red = new Color(154,1,0);

    public void init() throws IOException{
        idxImage = ImageIO.read(new File("asset/IDX.png"));
        kpeiImage = ImageIO.read(new File("asset/KPEI.jpg"));
        kseiImage = ImageIO.read(new File("asset/KSEI.png"));
        ciptadanaImage = ImageIO.read(new File("asset/ciptadana.png"));
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
                if(!(i == 1 | i == 2 | i == 4)){
                    totalWidth += sheet.getColumnWidthInPixels(i);
                }

            }
            totalWidth += 10;

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
            }

            int cellHeight = 30;
            int textPadding = 5;


            HashMap<String, HashMap<String, List<FixIncome>>> fixIncomeByCurrencyAndType =
                    fixIncomes.stream()
                            .collect(Collectors.groupingBy(
                                    FixIncome::getCurrency,
                                    HashMap::new,
                                    Collectors.groupingBy(FixIncome::getType, HashMap::new, Collectors.toList())
                            ));

            int cellColumn = 0;
            graphics.drawImage(ciptadanaImage, 5, 20, null);
            yPosition = ciptadanaImage.getHeight() + 20;

            for (String currency : fixIncomeByCurrencyAndType.keySet()) {

                Row headerRow = sheet.getRow(2);
                //String[] headers = new String[headerRow.getLastCellNum()];
                List<String> headers = new ArrayList<>();
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    headers.add(new DataFormatter().formatCellValue(headerRow.getCell(i)));
                }


                // 2. Draw the headers

                HashMap<String, List<FixIncome>> fixIncomeMap = fixIncomeByCurrencyAndType.get(currency);
                for(String type : fixIncomeMap.keySet()){
                    List<FixIncome> incomes = fixIncomeMap.get(type);
                    int xPositionHeader = 5;

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
                            graphics.setFont(new Font("Arial", Font.BOLD, 11).deriveFont(Font.BOLD | Font.ITALIC));
                            graphics.setColor(brown);
                            String title = generateTitle(type);
                            graphics.drawString(title, xPositionHeader + textPadding, yPosition + cellHeight - textPadding);


                            int titleWidth = graphics.getFontMetrics().stringWidth(title);

                            graphics.setColor(red);
                            graphics.drawString(" " + currency.toUpperCase(), xPositionHeader + textPadding + titleWidth, yPosition + cellHeight - textPadding);
                            yPosition += cellHeight;
                        }

                        if(!(j == 1 || j == 2 || j == 4)){
                            graphics.setFont(new Font("Arial", Font.PLAIN, 10));
                            graphics.setColor(Color.BLACK);
                            graphics.drawRect(xPositionHeader, yPosition, cellWidth, cellHeight);

                            graphics.setColor(red);
                            graphics.setFont(new Font("Arial", Font.BOLD, 10));
                            graphics.drawString(headers.get(j), xPositionHeader + textPadding, yPosition + cellHeight - textPadding);

                            xPositionHeader += cellWidth;
                        }
                    }
                    yPosition += cellHeight;



                    for (FixIncome income : incomes) {
                        int xPosition = 5;
                        for (int j = 0; j < 12; j++) { // 12 fields in FixIncome

                            if(!(j == 1 || j == 2 | j == 4)){
                                int cellWidth = (int) sheet.getColumnWidthInPixels(j);
                                graphics.setFont(new Font("Arial", Font.PLAIN, 10));
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
                    }
                   // yPosition += cellHeight;
                }
            }




            cellColumn = cellColumn / 3;

            graphics.setFont(new Font("Arial", Font.BOLD, 10).deriveFont(Font.BOLD | Font.ITALIC));
            graphics.setColor(Color.BLACK);
            graphics.drawString("** Prices are indicative, subject to availability and may change at any time.", textPadding, cellColumn + 5);
            graphics.drawImage(idxImage,  cellColumn - 550, cellColumn + 20, null);
            graphics.drawImage(kpeiImage, cellColumn - 400, cellColumn + 60, null);
            graphics.drawImage(kseiImage, cellColumn - 420, cellColumn + 20, null);

            graphics.setFont(new Font("Arial", Font.PLAIN, 10));
            graphics.setColor(red);
            graphics.drawString("PT Ciptadana Sekuritas Asia telah terdaftar dan diawasi oleh OJK", textPadding, cellColumn + 50);
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
