package com.ciptadana.uploadfixincome;

import jakarta.annotation.PostConstruct;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;

@Service
@Slf4j
public class IncomeService {
    BufferedImage idxImage;
    BufferedImage kpeiImage;
    BufferedImage kseiImage;
    BufferedImage ciptadanaImage;
    Color brown = new Color(154,102,51);
    Color red = new Color(154,1,0);
    Font ARIAL_PLAIN_11 = new Font("Arial", Font.PLAIN, 11);
    Font ARIAL_BOLD_ITALIC_12 = new Font("Arial", Font.BOLD, 12).deriveFont(Font.BOLD | Font.ITALIC);
    Font ARIAL_BOLD_10 = new Font("Arial", Font.BOLD, 10);
    Font ARIAL_BOLD_16 = new Font("Arial", Font.BOLD, 16);
    Font ARIAL_BOLD_18 = new Font("Arial", Font.BOLD, 18);
    Font ARIAL_BOLD_ITALIC_14 = new Font("Arial", Font.BOLD, 14).deriveFont(Font.BOLD | Font.ITALIC);
    SimpleDateFormat inputFormat = new SimpleDateFormat("dd/MM/yyyy");
    SimpleDateFormat outputFormat = new SimpleDateFormat("dd MMMM yyyy");

    @PostConstruct
    public void init() throws IOException{
        idxImage = ImageIO.read(new File("asset/IDX.png"));
        kpeiImage = ImageIO.read(new File("asset/KPEI.jpg"));
        kseiImage = ImageIO.read(new File("asset/KSEI.png"));
        ciptadanaImage = ImageIO.read(new File("asset/ciptadana.png"));
    }


    public BufferedImage processExcel(MultipartFile file) throws IOException {
        try (InputStream is = file.getInputStream()) {

            Workbook workbook;
            if (file.getOriginalFilename().endsWith(".xls")) {
                workbook = new HSSFWorkbook(is);
            } else if (file.getOriginalFilename().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(is);
            } else {
                throw new IllegalArgumentException("Invalid file type");
            }

            Sheet sheet = workbook.getSheetAt(0);

            int totalWidth = 40;
            short maxColumns = sheet.getRow(0).getLastCellNum();
            for (int i = 0; i < maxColumns; i++) {
                if(!(i == 1 | i == 2 | i == 4)){
                    totalWidth += sheet.getColumnWidthInPixels(i);
                }

            }


            int totalHeight = 0;
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                totalHeight += sheet.getRow(i).getHeight();
            }

            BufferedImage image = new BufferedImage(totalWidth, totalHeight, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = image.createGraphics();
            graphics.setColor(Color.WHITE);
            graphics.fillRect(0, 0, totalWidth, totalHeight);
            graphics.setFont(ARIAL_PLAIN_11);

            int yPosition = 0;
            List<FixIncome> fixIncomes = new ArrayList<>();

            Row dateRow = sheet.getRow(1);
            String date = formatDate(new DataFormatter().formatCellValue(dateRow.getCell(2)));

            for (int i = 3; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
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
            graphics.drawImage(ciptadanaImage, 8, 20, null);
            yPosition = ciptadanaImage.getHeight() + 60;


            graphics.setFont(ARIAL_BOLD_18);
            graphics.setColor(brown);
            graphics.drawString("INDIKASI BID - OFFER PT CIPTADANA SEKURITAS ASIA", totalWidth / 4, yPosition);
            yPosition += 20;

            graphics.setFont(ARIAL_BOLD_16);
            graphics.setColor(red);
            graphics.drawString(date, totalWidth / 2 - 40, yPosition);

            for (String currency : fixIncomeByCurrencyAndType.keySet()) {

                Row headerRow = sheet.getRow(2);
                List<String> headers = new ArrayList<>();
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    headers.add(new DataFormatter().formatCellValue(headerRow.getCell(i)));
                }


                HashMap<String, List<FixIncome>> fixIncomeMap = fixIncomeByCurrencyAndType.get(currency);
                for(String type : fixIncomeMap.keySet()){
                    List<FixIncome> incomes = fixIncomeMap.get(type);
                    int xPositionHeader = 10;
                    int totalCellWidth = - 209;
                    for (int j = 0; j < headers.size(); j++) {
                        cellColumn += (int) sheet.getColumnWidthInPixels(j);
                        totalCellWidth += (int) sheet.getColumnWidthInPixels(j);
                    }
                    yPosition += cellHeight;
                    for (int j = 0; j < headers.size(); j++) {

                        int cellWidth = (int) sheet.getColumnWidthInPixels(j);
                        if(j == 0){
                            graphics.setColor(Color.BLACK);
                            graphics.drawRect(xPositionHeader, yPosition, totalCellWidth, cellHeight);

                            graphics.setFont(ARIAL_BOLD_ITALIC_14);
                            graphics.setColor(brown);
                            String title = generateTitle(type);
                            graphics.drawString(title, xPositionHeader + textPadding, yPosition + cellHeight - textPadding);


                            int titleWidth = graphics.getFontMetrics().stringWidth(title);

                            graphics.setColor(brown);
                            graphics.drawString(" " + currency.toUpperCase(), xPositionHeader + textPadding + titleWidth, yPosition + cellHeight - textPadding);
                            yPosition += cellHeight;
                        }

                        if(!(j == 1 || j == 2 || j == 4)){
                            graphics.setFont(ARIAL_PLAIN_11);
                            graphics.setColor(Color.BLACK);
                            graphics.drawRect(xPositionHeader, yPosition, cellWidth, cellHeight);

                            graphics.setColor(red);

                            String data = headers.get(j);
                            FontMetrics fm = graphics.getFontMetrics();
                            int textWidth = fm.stringWidth(data);
                            int textHeight = fm.getHeight();

                            int centeredX = xPositionHeader + (cellWidth - textWidth) / 2;
                            int centeredY = yPosition + (cellHeight - textHeight) / 2 + fm.getAscent();

                            graphics.drawString(data, centeredX, centeredY);
                            xPositionHeader += cellWidth;
                        }
                    }
                    yPosition += cellHeight;



                    for (FixIncome income : incomes) {
                        int xPosition = 10;
                        for (int j = 0; j < 12; j++) {

                            if(!(j == 1 || j == 2 | j == 4)){
                                int cellWidth = (int) sheet.getColumnWidthInPixels(j);
                                  // fm.getAscent() gives the distance from the baseline to the top of the characters, which helps in vertically centering the text
                                graphics.setFont(ARIAL_PLAIN_11);
                                graphics.setColor(Color.BLACK);


                                String data = getFieldByIndex(income, j);

                                FontMetrics fm = graphics.getFontMetrics();
                                int textWidth = fm.stringWidth(data);
                                int textHeight = fm.getHeight();


                                int centeredX = xPosition + (cellWidth - textWidth) / 2;
                                int centeredY = yPosition + (cellHeight - textHeight) / 2 + fm.getAscent();
                                if(j == 0){
                                    centeredX = xPosition + (cellWidth - textWidth) / 4;
                                    centeredY = yPosition + (cellHeight - textHeight) / 2 + fm.getAscent();
                                }
                                graphics.drawRect(xPosition, yPosition, cellWidth, cellHeight);
                                graphics.drawString(data, centeredX, centeredY);

                                xPosition += cellWidth;
                            }
                        }
                        yPosition += cellHeight;
                    }
                }
            }


            yPosition += 30;
            graphics.setFont(ARIAL_BOLD_ITALIC_12);
            graphics.setColor(Color.BLACK);
            graphics.drawString("   ** Prices are indicative, subject to availability and may change at any time.", textPadding, yPosition);

            graphics.drawImage(idxImage,  totalWidth / 2 + 10, yPosition, null);
            graphics.drawImage(kpeiImage, totalWidth / 2 + 80, yPosition, null);
            graphics.drawImage(kseiImage, totalWidth / 2 + 200, yPosition, null);

            yPosition += 30;
            graphics.setFont(ARIAL_PLAIN_11);
            graphics.setColor(red);
            graphics.drawString("   PT Ciptadana Sekuritas Asia telah terdaftar dan diawasi oleh OJK", textPadding, yPosition);
            graphics.dispose();
            image = trim(image, Color.white);
           // ImageIO.write(image, "jpeg", new File(imageOutputPath));
            return image;

        }


    }

    public BufferedImage trim(BufferedImage image, Color bgColor) {
        int top = 0, left = 0, bottom = image.getHeight() - 1, right = image.getWidth() - 1;

        while (top < bottom && isRowEmpty(image, top, bgColor)) {
            top++;
        }
        while (bottom > top && isRowEmpty(image, bottom, bgColor)) {
            bottom--;
        }
        while (left < right && isColumnEmpty(image, left, bgColor)) {
            left++;
        }
        while (right > left && isColumnEmpty(image, right, bgColor)) {
            right--;
        }

        return image.getSubimage(left - 5, top - 10, right - left + 15, bottom - top + 15);
    }

    private boolean isRowEmpty(BufferedImage image, int row, Color bgColor) {
        int width = image.getWidth();
        for (int x = 0; x < width; x++) {
            if (!new Color(image.getRGB(x, row)).equals(bgColor)) {
                return false;
            }
        }
        return true;
    }

    private boolean isColumnEmpty(BufferedImage image, int column, Color bgColor) {
        int height = image.getHeight();
        for (int y = 0; y < height; y++) {
            if (!new Color(image.getRGB(column, y)).equals(bgColor)) {
                return false;
            }
        }
        return true;
    }

    private String formatDate(String inputDate) {
        try{
            Date date = inputFormat.parse(inputDate);
            return outputFormat.format(date).toUpperCase();
        }catch (Exception e){
            log.error(e.getMessage());
        }
        return "INVALID FORMAT";
    }

    private String getFieldByIndex(FixIncome income, int index) {
        switch(index) {
            case 0: return formatIssuerName(income.getIssuerName());
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

    public String formatIssuerName(String input) {
        if (input != null && input.startsWith("FR")) {
            String number = input.substring(2).replaceFirst("^0+", "");
            return "Fixed Rate " + number;
        }
        return input;
    }
}
