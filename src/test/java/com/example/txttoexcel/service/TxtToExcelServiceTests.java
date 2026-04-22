package com.example.txttoexcel.service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class TxtToExcelServiceTests {

    private final TxtToExcelService txtToExcelService = new TxtToExcelService();

    @Test
    void convertsDelimitedTxtIntoStyledWorkbook() throws IOException {
        String source = """
                Name|Department|Score
                Alice|Sales|98
                Bob|Finance|87
                """;

        byte[] workbookBytes = txtToExcelService.convert(
                new ByteArrayInputStream(source.getBytes(StandardCharsets.UTF_8)),
                "staff.txt");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(workbookBytes))) {
            Sheet sheet = workbook.getSheet("TXT Data");
            Sheet summarySheet = workbook.getSheet("Summary");

            assertThat(sheet).isNotNull();
            assertThat(summarySheet).isNotNull();
            assertThat(summarySheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Conversion Summary");
            assertThat(summarySheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("staff.txt");
            assertThat(summarySheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Pipe separated");
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Name");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Department");
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("Alice");
            assertThat(sheet.getRow(1).getCell(2).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(sheet.getRow(1).getCell(2).getNumericCellValue()).isEqualTo(98.0d);
        }
    }

    @Test
    void convertsIndentedVisanetReportIntoStructuredRows() throws IOException {
        String source = """
                \fREPORT ID: SMS601C                                 VISANET INTEGRATED PAYMENT SYSTEM                 PAGE NUMBER         :        1
                                                                                                   FUNDS XFR: 1000108517 CAMBODIA ASIA B                  SINGLECONNECT / VISA                          ONLINE SETTLMNT DATE:  26MAR26
                                                                                                   PROCESSOR: 4520580002 CAMBODIA ASIA BANK            ACQUIRER TRANSACTION DETAIL                      REPORT DATE         :  26MAR26
                                                                                                   AFFILIATE: 4520580002 CAMBODIA ASIA BANK                BY CARDHOLDER NUMBER                         REPORT TIME         : 10:37:20
                                                                                                   SRE      : 9000320257 CAB ATM NW 2                                                                   VSS PROCESSING DATE :  26MAR26

                                                                                                   -----------------------------------------------------------------------------------------------------------------------------------
                                                                                                   BAT XMIT(GMT)/LOCL                     RETRIEVAL    TRACE  ISSUER ID/  TRAN PROCSS ENT REAS CN/ RSP  --TRANSACTION--   SETTLEMENT
                                                                                                   NUM DATE  TIME     CARD NUMBER         REF NUMBER   NUMBER TRMNL/NAME  TYPE CODE   MOD CODE STP CD        AMOUNT CUR   AMOUNT (USD)
                                                                                                   -----------------------------------------------------------------------------------------------------------------------------------

                                                                                                    02 25MAR 10:51:51 4005360035804595   608410136350 136350 400536    0200 011000 051      02  00      2,060.00 USD     2,060.00CR
                                                                                                                      CA ID: 10008219  10013860               25/NAGA1 VIP Gflr T70       /KH  0000                  FPI: 8C0
                                                                                                                                                       ATC: 00005                     CI: 1
                                                                                                                      TR ID: 466084391114279    ACI: E VC: HR2F                                  SCHG:        60.00
                                                                                                       FEE JURIS: VISA INTERNATIONAL       ROUTING: 1 A.P.      -C.E.M.E.A       FEE LEVEL: ATM AF                          3.000000CR

                                                                                                    22 25MAR 20:55:34 4005360035804595    608420248160 248160 400536      0200 011000 051      02  00      1,030.00 USD     1,030.00CR
                                                                                                                      CA ID: 10008191  10013832               25/Casino South Tower Gflr T/KH  0000                  FPI: 8C0
                                                                                                                                                       ATC: 00006                     CI: 1
                                                                                                                      TR ID: 386084753341462    ACI: E VC: 9S9K
                """;

        byte[] workbookBytes = txtToExcelService.convert(
                new ByteArrayInputStream(source.getBytes(StandardCharsets.UTF_8)),
                "visanet-report.txt");

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(workbookBytes))) {
            Sheet sheet = workbook.getSheet("TXT Data");
            Sheet summarySheet = workbook.getSheet("Summary");

            assertThat(summarySheet.getRow(2).getCell(1).getStringCellValue())
                    .isEqualTo("Visanet fixed-width transaction report");
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue())
                    .isEqualTo("VISANET INTEGRATED PAYMENT SYSTEM");
            assertThat(sheet.getRow(0).getCell(0).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue())
                    .contains("ACQUIRER TRANSACTION DETAIL")
                    .contains("BY CARDHOLDER NUMBER")
                    .contains("Report ID: SMS601C");
            assertThat(sheet.getRow(1).getCell(0).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
            assertThat(findSummaryValue(summarySheet, "Report ID")).isEqualTo("SMS601C");
            assertThat(findSummaryValue(summarySheet, "Funds XFR")).isEqualTo("1000108517 CAMBODIA ASIA B");
            assertThat(findSummaryValue(summarySheet, "Report time")).isEqualTo("10:37:20");
            assertThat(findSummaryValue(summarySheet, "Total amount")).isEqualTo("3,090.00");
            assertThat(findSummaryValue(summarySheet, "Total surcharge")).isEqualTo("60.00");
            assertThat(findSummaryValue(summarySheet, "Total fee")).isEqualTo("3");

            int headerRowIndex = findHeaderRowIndex(sheet, "Batch No");
            int batchColumn = findColumnIndex(sheet, headerRowIndex, "Batch No");
            int cardNumberColumn = findColumnIndex(sheet, headerRowIndex, "Card Number");
            int amountColumn = findColumnIndex(sheet, headerRowIndex, "Amount");
            int terminalColumn = findColumnIndex(sheet, headerRowIndex, "Terminal/Name");
            int atcColumn = findColumnIndex(sheet, headerRowIndex, "ATC");
            int transactionIdColumn = findColumnIndex(sheet, headerRowIndex, "Transaction ID");
            int vcColumn = findColumnIndex(sheet, headerRowIndex, "VC");
            int ciColumn = findColumnIndex(sheet, headerRowIndex, "CI");
            int surchargeColumn = findColumnIndex(sheet, headerRowIndex, "Surcharge");
            int unparsedDetailColumn = findColumnIndex(sheet, headerRowIndex, "Unparsed Detail");

            assertThat(sheet.getRow(headerRowIndex + 1).getCell(batchColumn).getStringCellValue()).isEqualTo("02");
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(cardNumberColumn).getCellType()).isEqualTo(CellType.STRING);
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(cardNumberColumn).getStringCellValue()).isEqualTo("4005360035804595");
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(amountColumn).getNumericCellValue()).isEqualTo(2060.0d);
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(terminalColumn).getStringCellValue()).isEqualTo("25/NAGA1 VIP Gflr T70");
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(atcColumn).getStringCellValue()).isEqualTo("00005");
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(transactionIdColumn).getStringCellValue()).isEqualTo("466084391114279");
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(ciColumn).getStringCellValue()).isEqualTo("1");
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(surchargeColumn).getNumericCellValue()).isEqualTo(60.0d);
            assertThat(sheet.getRow(headerRowIndex + 1).getCell(unparsedDetailColumn).getCellType()).isEqualTo(CellType.BLANK);

            assertThat(sheet.getRow(headerRowIndex + 2).getCell(batchColumn).getStringCellValue()).isEqualTo("22");
            assertThat(sheet.getRow(headerRowIndex + 2).getCell(atcColumn).getStringCellValue()).isEqualTo("00006");
            assertThat(sheet.getRow(headerRowIndex + 2).getCell(transactionIdColumn).getStringCellValue()).isEqualTo("386084753341462");
            assertThat(sheet.getRow(headerRowIndex + 2).getCell(vcColumn).getStringCellValue()).isEqualTo("9S9K");
            assertThat(sheet.getRow(headerRowIndex + 2).getCell(ciColumn).getStringCellValue()).isEqualTo("1");
            assertThat(sheet.getRow(headerRowIndex + 2).getCell(unparsedDetailColumn).getCellType()).isEqualTo(CellType.BLANK);
        }
    }

    private int findHeaderRowIndex(Sheet sheet, String headerName) {
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            if (sheet.getRow(rowIndex) == null) {
                continue;
            }

            for (int columnIndex = 0; columnIndex < sheet.getRow(rowIndex).getLastCellNum(); columnIndex++) {
                if (sheet.getRow(rowIndex).getCell(columnIndex) != null
                        && headerName.equals(sheet.getRow(rowIndex).getCell(columnIndex).getStringCellValue())) {
                    return rowIndex;
                }
            }
        }

        throw new AssertionError("Header row not found for: " + headerName);
    }

    private int findColumnIndex(Sheet sheet, int headerRowIndex, String headerName) {
        for (int columnIndex = 0; columnIndex < sheet.getRow(headerRowIndex).getLastCellNum(); columnIndex++) {
            if (headerName.equals(sheet.getRow(headerRowIndex).getCell(columnIndex).getStringCellValue())) {
                return columnIndex;
            }
        }

        throw new AssertionError("Header not found: " + headerName);
    }

    private String findSummaryValue(Sheet summarySheet, String label) {
        for (int rowIndex = 0; rowIndex <= summarySheet.getLastRowNum(); rowIndex++) {
            if (summarySheet.getRow(rowIndex) == null || summarySheet.getRow(rowIndex).getCell(0) == null) {
                continue;
            }

            if (label.equals(summarySheet.getRow(rowIndex).getCell(0).getStringCellValue())) {
                return summarySheet.getRow(rowIndex).getCell(1).getStringCellValue();
            }
        }

        throw new AssertionError("Summary label not found: " + label);
    }
}
