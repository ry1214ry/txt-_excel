package com.example.txttoexcel.service;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

@Service
public class TxtToExcelService {

    private static final String DATA_SHEET_NAME = "TXT Data";
    private static final String SUMMARY_SHEET_NAME = "Summary";
    private static final List<String> DELIMITER_CANDIDATES = List.of("\t", "|", ";", ",");
    private static final DateTimeFormatter EXPORT_TIME_FORMAT =
            DateTimeFormatter.ofPattern("dd MMM uuuu HH:mm");
    private static final Pattern REPORT_DIVIDER_PATTERN = Pattern.compile("^-{10,}\\s*$");
    private static final Pattern VISANET_ROW_START_PATTERN = Pattern.compile(
            "^\\s*(?<batch>\\d{2})\\s+(?<date>\\d{2}[A-Z]{3})\\s+(?<time>\\d{2}:\\d{2}:\\d{2})\\s+"
                    + "(?<card>\\d{13,19})\\s+(?<retrieval>\\d+)\\s+(?<trace>\\d+)\\s+(?<issuer>\\d+)\\s+"
                    + "(?<tranType>\\d{4})\\s+(?<processCode>\\d{6})\\s+(?<entryMode>\\d{3})\\s+"
                    + "(?<reasonCode>\\d{2})\\s+(?<rspCode>\\d{2})\\s+(?<amount>[\\d,]+\\.\\d{2})\\s+"
                    + "(?<currency>[A-Z]{3})\\s+(?<settlement>[\\d,]+\\.\\d{2}(?:CR|DR)?)\\s*$");
    private static final Pattern VISANET_CA_LINE_PATTERN = Pattern.compile(
            "^\\s*CA ID:\\s*(?<caId>\\d+)\\s+(?<caRef>\\d+)(?<details>.*?)(?:\\s+FPI:\\s*(?<fpi>\\S+))?\\s*$");
    private static final Pattern VISANET_LOCATION_PATTERN = Pattern.compile(
            "^(?<terminal>.+?)\\s+(?<country>/[A-Z]{2})\\s+(?<stpCode>\\d{4})\\s*$");
    private static final Pattern VISANET_TRANSACTION_LINE_PATTERN = Pattern.compile(
            "^\\s*(?:CD SQ:\\s*(?<cardSequence>\\S+)\\s+)?TR ID:\\s*(?<transactionId>\\d+)\\s+ACI:\\s*(?<aci>\\S+)\\s+"
                    + "VC:\\s*(?<vc>\\S+).*$");
    private static final Pattern VISANET_FEE_LINE_PATTERN = Pattern.compile(
            "^\\s*FEE JURIS:\\s*(?<feeJuris>.+?)\\s+ROUTING:\\s*(?<routing>.+?)\\s+FEE LEVEL:\\s*(?<feeLevel>.+?)\\s+"
                    + "(?<feeAmount>[\\d,]+\\.\\d+(?:CR|DR)?)\\s*$");
    private static final Pattern VISANET_ATC_PATTERN = Pattern.compile("\\bATC:\\s*(?<atc>\\S+)");
    private static final Pattern VISANET_CI_PATTERN = Pattern.compile("\\bCI:\\s*(?<ci>\\S+)");
    private static final Pattern VISANET_SURCHARGE_PATTERN = Pattern.compile(
            "\\bSCHG:\\s*(?<surcharge>[\\d,]+\\.\\d{2})");
    private static final Pattern NUMERIC_VALUE_PATTERN = Pattern.compile(
            "^-?(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?$");
    private static final List<String> VISANET_METADATA_LABELS = List.of(
            "REPORT ID",
            "PAGE NUMBER",
            "FUNDS XFR",
            "ONLINE SETTLMNT DATE",
            "PROCESSOR",
            "REPORT DATE",
            "AFFILIATE",
            "REPORT TIME",
            "SRE",
            "VSS PROCESSING DATE");
    private static final List<String> VISANET_REPORT_TITLES = List.of(
            "VISANET INTEGRATED PAYMENT SYSTEM",
            "SINGLECONNECT / VISA",
            "ACQUIRER TRANSACTION DETAIL",
            "BY CARDHOLDER NUMBER");
    private static final List<String> VISANET_HEADERS = List.of(
            "Batch No",
            "Date",
            "Time",
            "Card Number",
            "Retrieval Ref",
            "Trace Number",
            "Issuer ID",
            "Tran Type",
            "Process Code",
            "Entry Mode",
            "Reason Code",
            "Rsp Code",
            "Amount",
            "Currency",
            "Settlement (USD)",
            "CA ID",
            "CA Ref",
            "Terminal/Name",
            "Country",
            "STP Code",
            "FPI",
            "ATC",
            "Card Seq",
            "Transaction ID",
            "ACI",
            "VC",
            "CI",
            "Surcharge",
            "Fee Jurisdiction",
            "Routing",
            "Fee Level",
            "Fee Amount",
            "Unparsed Detail");
    private static final int VISANET_CARD_NUMBER_INDEX = VISANET_HEADERS.indexOf("Card Number");
    private static final int VISANET_AMOUNT_INDEX = VISANET_HEADERS.indexOf("Amount");
    private static final int VISANET_SETTLEMENT_INDEX = VISANET_HEADERS.indexOf("Settlement (USD)");
    private static final int VISANET_SURCHARGE_INDEX = VISANET_HEADERS.indexOf("Surcharge");
    private static final int VISANET_FEE_AMOUNT_INDEX = VISANET_HEADERS.indexOf("Fee Amount");
    private static final Set<Integer> VISANET_TEXT_ONLY_COLUMNS = Set.of(
            0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 24, 25, 26);
    private static final int VISANET_FREEZE_COLUMNS = 4;

    public byte[] convert(InputStream inputStream, String originalFilename) throws IOException {
        ParsedTxtData parsedTxtData = parseTxtFile(inputStream, originalFilename);

        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Sheet summarySheet = workbook.createSheet(SUMMARY_SHEET_NAME);
            Sheet sheet = workbook.createSheet(DATA_SHEET_NAME);
            Styles styles = createStyles(workbook);
            int lastColumnIndex = Math.max(parsedTxtData.headers().size() - 1, 0);
            int headerRowIndex = createDataSheetIntro(sheet, parsedTxtData, styles);

            createSummarySheet(summarySheet, parsedTxtData, styles);
            createHeaderRow(sheet, headerRowIndex, parsedTxtData.headers(), styles.header());
            createDataRows(
                    sheet,
                    headerRowIndex + 1,
                    parsedTxtData.rows(),
                    parsedTxtData.headers().size(),
                    parsedTxtData.textOnlyColumns(),
                    styles);

            workbook.setSheetOrder(DATA_SHEET_NAME, 0);
            workbook.setActiveSheet(0);
            workbook.setFirstVisibleTab(0);
            sheet.createFreezePane(parsedTxtData.frozenColumnCount(), headerRowIndex + 1);
            if (!parsedTxtData.rows().isEmpty()) {
                sheet.setAutoFilter(new CellRangeAddress(
                        headerRowIndex,
                        headerRowIndex + parsedTxtData.rows().size(),
                        0,
                        lastColumnIndex));
            }

            sizeDataColumns(sheet, parsedTxtData.headers());

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private ParsedTxtData parseTxtFile(InputStream inputStream, String originalFilename) throws IOException {
        List<String> nonBlankLines = readNonBlankLines(inputStream);
        if (nonBlankLines.isEmpty()) {
            throw new IllegalArgumentException("The uploaded TXT file is empty.");
        }

        String cleanFileName = StringUtils.hasText(originalFilename) ? originalFilename : "uploaded.txt";
        if (looksLikeVisanetReport(nonBlankLines)) {
            ParsedTxtData reportData = parseVisanetReport(nonBlankLines, cleanFileName);
            if (!reportData.rows().isEmpty()) {
                return reportData;
            }
        }

        String detectedDelimiter = detectDelimiter(nonBlankLines);

        if (detectedDelimiter == null) {
            List<List<String>> rows = nonBlankLines.stream()
                    .map(line -> List.of(line.strip()))
                    .toList();
            return new ParsedTxtData(
                    List.of("Line"),
                    rows,
                    "Single-column text",
                    cleanFileName,
                    Set.of(),
                    0,
                    Map.of(),
                    "",
                    "");
        }

        List<List<String>> splitRows = nonBlankLines.stream()
                .map(line -> splitLine(line, detectedDelimiter))
                .toList();
        int maxColumns = splitRows.stream()
                .mapToInt(List::size)
                .max()
                .orElse(1);

        List<String> headers;
        List<List<String>> rows;
        if (splitRows.size() > 1) {
            headers = normalizeHeaders(splitRows.getFirst(), maxColumns);
            rows = splitRows.stream()
                    .skip(1)
                    .map(row -> padRow(row, maxColumns))
                    .toList();
        } else {
            headers = createGenericHeaders(maxColumns);
            rows = splitRows.stream()
                    .map(row -> padRow(row, maxColumns))
                    .toList();
        }

        return new ParsedTxtData(
                headers,
                rows,
                describeDelimiter(detectedDelimiter),
                cleanFileName,
                Set.of(),
                0,
                Map.of(),
                "",
                "");
    }

    private boolean looksLikeVisanetReport(List<String> lines) {
        boolean hasVisanetHeader = lines.stream()
                .anyMatch(line -> line.contains("VISANET INTEGRATED PAYMENT SYSTEM"));
        boolean hasTransactionRows = lines.stream()
                .anyMatch(line -> VISANET_ROW_START_PATTERN.matcher(line).matches());
        return hasVisanetHeader && hasTransactionRows;
    }

    private ParsedTxtData parseVisanetReport(List<String> lines, String originalFilename) {
        List<List<String>> rows = new ArrayList<>();
        Map<String, String> currentRow = null;
        Map<String, String> visanetMetadata = extractVisanetMetadata(lines);

        for (String line : lines) {
            if (isSkippableVisanetLine(line)) {
                continue;
            }

            Matcher startMatcher = VISANET_ROW_START_PATTERN.matcher(line);
            if (startMatcher.matches()) {
                if (currentRow != null) {
                    rows.add(buildVisanetRow(currentRow));
                }
                currentRow = createVisanetRowMap();
                populate(currentRow, "Batch No", startMatcher.group("batch"));
                populate(currentRow, "Date", startMatcher.group("date"));
                populate(currentRow, "Time", startMatcher.group("time"));
                populate(currentRow, "Card Number", startMatcher.group("card"));
                populate(currentRow, "Retrieval Ref", startMatcher.group("retrieval"));
                populate(currentRow, "Trace Number", startMatcher.group("trace"));
                populate(currentRow, "Issuer ID", startMatcher.group("issuer"));
                populate(currentRow, "Tran Type", startMatcher.group("tranType"));
                populate(currentRow, "Process Code", startMatcher.group("processCode"));
                populate(currentRow, "Entry Mode", startMatcher.group("entryMode"));
                populate(currentRow, "Reason Code", startMatcher.group("reasonCode"));
                populate(currentRow, "Rsp Code", startMatcher.group("rspCode"));
                populate(currentRow, "Amount", startMatcher.group("amount"));
                populate(currentRow, "Currency", startMatcher.group("currency"));
                populate(currentRow, "Settlement (USD)", startMatcher.group("settlement"));
                continue;
            }

            if (currentRow == null) {
                continue;
            }

            if (populateFromCaLine(currentRow, line)) {
                continue;
            }

            if (populateFromTransactionLine(currentRow, line)) {
                populateInlineVisanetMarkers(currentRow, line);
                continue;
            }

            if (populateFromFeeLine(currentRow, line)) {
                continue;
            }

            if (populateInlineVisanetMarkers(currentRow, line)) {
                continue;
            }

            appendUnparsedDetail(currentRow, line);
        }

        if (currentRow != null) {
            rows.add(buildVisanetRow(currentRow));
        }

        return new ParsedTxtData(
                VISANET_HEADERS,
                rows,
                "Visanet fixed-width transaction report",
                originalFilename,
                VISANET_TEXT_ONLY_COLUMNS,
                VISANET_FREEZE_COLUMNS,
                buildVisanetSummaryDetails(visanetMetadata, rows),
                buildVisanetSheetTitle(lines),
                buildVisanetSheetSubtitle(lines, visanetMetadata));
    }

    private List<String> readNonBlankLines(InputStream inputStream) throws IOException {
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8))) {
            return reader.lines()
                    .map(this::normalizeLine)
                    .filter(line -> !line.isBlank())
                    .toList();
        }
    }

    private String normalizeLine(String line) {
        return line
                .replace("\uFEFF", "")
                .replace('\f', ' ')
                .stripTrailing();
    }

    private String detectDelimiter(List<String> lines) {
        String bestDelimiter = null;
        int bestMatchedLines = 0;
        int bestAverageColumns = 0;

        for (String candidate : DELIMITER_CANDIDATES) {
            int matchedLines = 0;
            int totalColumns = 0;

            for (String line : lines) {
                int columnCount = splitLine(line, candidate).size();
                if (columnCount > 1) {
                    matchedLines++;
                    totalColumns += columnCount;
                }
            }

            int averageColumns = matchedLines == 0 ? 0 : totalColumns / matchedLines;
            boolean betterMatch = matchedLines > bestMatchedLines
                    || (matchedLines == bestMatchedLines && averageColumns > bestAverageColumns);
            if (betterMatch && matchedLines > 0) {
                bestDelimiter = candidate;
                bestMatchedLines = matchedLines;
                bestAverageColumns = averageColumns;
            }
        }

        return bestDelimiter;
    }

    private List<String> splitLine(String line, String delimiter) {
        return List.of(line.split(Pattern.quote(delimiter), -1)).stream()
                .map(String::trim)
                .toList();
    }

    private Map<String, String> createVisanetRowMap() {
        Map<String, String> values = new LinkedHashMap<>();
        for (String header : VISANET_HEADERS) {
            values.put(header, "");
        }
        return values;
    }

    private boolean populateFromCaLine(Map<String, String> rowValues, String line) {
        Matcher matcher = VISANET_CA_LINE_PATTERN.matcher(line);
        if (!matcher.matches()) {
            return false;
        }

        populate(rowValues, "CA ID", matcher.group("caId"));
        populate(rowValues, "CA Ref", matcher.group("caRef"));
        populate(rowValues, "FPI", matcher.group("fpi"));

        String details = matcher.group("details") == null ? "" : matcher.group("details").trim();
        if (!details.isEmpty()) {
            Matcher locationMatcher = VISANET_LOCATION_PATTERN.matcher(details);
            if (locationMatcher.matches()) {
                populate(rowValues, "Terminal/Name", locationMatcher.group("terminal"));
                populate(rowValues, "Country", locationMatcher.group("country"));
                populate(rowValues, "STP Code", locationMatcher.group("stpCode"));
            } else {
                populate(rowValues, "Terminal/Name", details);
            }
        }

        return true;
    }

    private boolean populateFromTransactionLine(Map<String, String> rowValues, String line) {
        Matcher matcher = VISANET_TRANSACTION_LINE_PATTERN.matcher(line);
        if (!matcher.matches()) {
            return false;
        }

        populate(rowValues, "Card Seq", matcher.group("cardSequence"));
        populate(rowValues, "Transaction ID", matcher.group("transactionId"));
        populate(rowValues, "ACI", matcher.group("aci"));
        populate(rowValues, "VC", matcher.group("vc"));
        return true;
    }

    private boolean populateInlineVisanetMarkers(Map<String, String> rowValues, String line) {
        boolean matched = false;
        matched |= populateFromInlineMatcher(rowValues, line, VISANET_ATC_PATTERN, "ATC", "atc");
        matched |= populateFromInlineMatcher(rowValues, line, VISANET_CI_PATTERN, "CI", "ci");
        matched |= populateFromInlineMatcher(rowValues, line, VISANET_SURCHARGE_PATTERN, "Surcharge", "surcharge");
        return matched;
    }

    private boolean populateFromInlineMatcher(
            Map<String, String> rowValues,
            String line,
            Pattern pattern,
            String key,
            String groupName) {
        Matcher matcher = pattern.matcher(line);
        if (!matcher.find()) {
            return false;
        }

        populate(rowValues, key, matcher.group(groupName));
        return true;
    }

    private boolean populateFromFeeLine(Map<String, String> rowValues, String line) {
        Matcher matcher = VISANET_FEE_LINE_PATTERN.matcher(line);
        if (!matcher.matches()) {
            return false;
        }

        populate(rowValues, "Fee Jurisdiction", matcher.group("feeJuris"));
        populate(rowValues, "Routing", matcher.group("routing"));
        populate(rowValues, "Fee Level", matcher.group("feeLevel"));
        populate(rowValues, "Fee Amount", matcher.group("feeAmount"));
        return true;
    }

    private void populate(Map<String, String> rowValues, String key, String value) {
        if (value != null && !value.isBlank()) {
            rowValues.put(key, value.trim());
        }
    }

    private void appendUnparsedDetail(Map<String, String> rowValues, String line) {
        String cleanLine = line.strip();
        if (cleanLine.isEmpty()) {
            return;
        }

        String existingValue = rowValues.get("Unparsed Detail");
        rowValues.put("Unparsed Detail", existingValue.isEmpty() ? cleanLine : existingValue + " | " + cleanLine);
    }

    private List<String> buildVisanetRow(Map<String, String> rowValues) {
        return VISANET_HEADERS.stream()
                .map(header -> rowValues.getOrDefault(header, ""))
                .toList();
    }

    private boolean isSkippableVisanetLine(String line) {
        String normalizedLine = line.strip();
        return REPORT_DIVIDER_PATTERN.matcher(normalizedLine).matches()
                || normalizedLine.startsWith("BAT XMIT")
                || normalizedLine.startsWith("NUM DATE")
                || normalizedLine.startsWith("REPORT ID:")
                || normalizedLine.startsWith("FUNDS XFR:")
                || normalizedLine.startsWith("PROCESSOR:")
                || normalizedLine.startsWith("AFFILIATE:")
                || normalizedLine.startsWith("SRE")
                || normalizedLine.contains("VISANET INTEGRATED PAYMENT SYSTEM")
                || normalizedLine.contains("SINGLECONNECT / VISA")
                || normalizedLine.contains("ACQUIRER TRANSACTION DETAIL")
                || normalizedLine.contains("BY CARDHOLDER NUMBER");
    }

    private Map<String, String> buildVisanetSummaryDetails(Map<String, String> visanetMetadata, List<List<String>> rows) {
        LinkedHashMap<String, String> summaryDetails = new LinkedHashMap<>();
        summaryDetails.putAll(visanetMetadata);

        if (!rows.isEmpty()) {
            summaryDetails.put("Unique cards", String.valueOf(countUniqueNonBlankValues(rows, VISANET_CARD_NUMBER_INDEX)));
            summaryDetails.put("Total amount", formatSummaryNumber(sumColumn(rows, VISANET_AMOUNT_INDEX), "#,##0.00"));
            summaryDetails.put(
                    "Total settlement (USD)",
                    formatSummaryNumber(sumColumn(rows, VISANET_SETTLEMENT_INDEX), "#,##0.00"));
            summaryDetails.put(
                    "Total surcharge",
                    formatSummaryNumber(sumColumn(rows, VISANET_SURCHARGE_INDEX), "#,##0.00"));
            summaryDetails.put("Total fee", formatSummaryNumber(sumColumn(rows, VISANET_FEE_AMOUNT_INDEX), "#,##0.######"));
        }

        return Collections.unmodifiableMap(new LinkedHashMap<>(summaryDetails));
    }

    private String buildVisanetSheetTitle(List<String> lines) {
        for (String line : lines) {
            if (line.contains("VISANET INTEGRATED PAYMENT SYSTEM")) {
                return "VISANET INTEGRATED PAYMENT SYSTEM";
            }
        }

        return "Transaction Report";
    }

    private String buildVisanetSheetSubtitle(List<String> lines, Map<String, String> visanetMetadata) {
        List<String> parts = new ArrayList<>();

        if (lines.stream().anyMatch(line -> line.contains("ACQUIRER TRANSACTION DETAIL"))) {
            parts.add("ACQUIRER TRANSACTION DETAIL");
        }
        if (lines.stream().anyMatch(line -> line.contains("BY CARDHOLDER NUMBER"))) {
            parts.add("BY CARDHOLDER NUMBER");
        }
        if (StringUtils.hasText(visanetMetadata.get("Report ID"))) {
            parts.add("Report ID: " + visanetMetadata.get("Report ID"));
        }
        if (StringUtils.hasText(visanetMetadata.get("Report date"))) {
            parts.add("Report date: " + visanetMetadata.get("Report date"));
        }
        if (StringUtils.hasText(visanetMetadata.get("Report time"))) {
            parts.add("Report time: " + visanetMetadata.get("Report time"));
        }

        return String.join(" | ", parts);
    }

    private Map<String, String> extractVisanetMetadata(List<String> lines) {
        LinkedHashMap<String, String> metadata = new LinkedHashMap<>();
        addMetadataIfPresent(metadata, "Report ID", extractFirstMetadataValue(lines, "REPORT ID"));
        addMetadataIfPresent(metadata, "Page Number", extractFirstMetadataValue(lines, "PAGE NUMBER"));
        addMetadataIfPresent(metadata, "Funds XFR", extractFirstMetadataValue(lines, "FUNDS XFR"));
        addMetadataIfPresent(metadata, "Processor", extractFirstMetadataValue(lines, "PROCESSOR"));
        addMetadataIfPresent(metadata, "Affiliate", extractFirstMetadataValue(lines, "AFFILIATE"));
        addMetadataIfPresent(metadata, "SRE", extractFirstMetadataValue(lines, "SRE"));
        addMetadataIfPresent(
                metadata,
                "Online settlement date",
                extractFirstMetadataValue(lines, "ONLINE SETTLMNT DATE"));
        addMetadataIfPresent(metadata, "Report date", extractFirstMetadataValue(lines, "REPORT DATE"));
        addMetadataIfPresent(metadata, "Report time", extractFirstMetadataValue(lines, "REPORT TIME"));
        addMetadataIfPresent(
                metadata,
                "VSS processing date",
                extractFirstMetadataValue(lines, "VSS PROCESSING DATE"));
        return metadata;
    }

    private void addMetadataIfPresent(Map<String, String> metadata, String label, String value) {
        if (StringUtils.hasText(value)) {
            metadata.put(label, value);
        }
    }

    private String extractFirstMetadataValue(List<String> lines, String label) {
        for (String line : lines) {
            String value = extractMetadataValue(line, label);
            if (StringUtils.hasText(value)) {
                return value;
            }
        }

        return "";
    }

    private String extractMetadataValue(String line, String label) {
        LabelMatch currentLabelMatch = findLabelMatch(line, label, 0);
        if (currentLabelMatch == null) {
            return "";
        }

        int nextLabelStart = -1;
        for (String candidateLabel : VISANET_METADATA_LABELS) {
            if (candidateLabel.equals(label)) {
                continue;
            }

            LabelMatch candidateMatch = findLabelMatch(line, candidateLabel, currentLabelMatch.end());
            if (candidateMatch != null && (nextLabelStart == -1 || candidateMatch.start() < nextLabelStart)) {
                nextLabelStart = candidateMatch.start();
            }
        }

        String valueSection = nextLabelStart >= 0
                ? line.substring(currentLabelMatch.end(), nextLabelStart)
                : line.substring(currentLabelMatch.end());
        String cleanValue = valueSection;
        for (String reportTitle : VISANET_REPORT_TITLES) {
            cleanValue = cleanValue.replace(reportTitle, " ");
        }

        return cleanValue.replaceAll("\\s{2,}", " ").trim();
    }

    private LabelMatch findLabelMatch(String line, String label, int startIndex) {
        Pattern pattern = Pattern.compile("(?<!\\S)" + Pattern.quote(label) + "\\s*:");
        Matcher matcher = pattern.matcher(line);
        if (!matcher.find(startIndex)) {
            return null;
        }

        return new LabelMatch(matcher.start(), matcher.end());
    }

    private long countUniqueNonBlankValues(List<List<String>> rows, int columnIndex) {
        return rows.stream()
                .map(row -> columnIndex < row.size() ? row.get(columnIndex).trim() : "")
                .filter(StringUtils::hasText)
                .distinct()
                .count();
    }

    private BigDecimal sumColumn(List<List<String>> rows, int columnIndex) {
        BigDecimal total = BigDecimal.ZERO;

        for (List<String> row : rows) {
            if (columnIndex >= row.size()) {
                continue;
            }

            BigDecimal value = parseSignedNumericValue(row.get(columnIndex));
            if (value != null) {
                total = total.add(value);
            }
        }

        return total;
    }

    private String formatSummaryNumber(BigDecimal value, String pattern) {
        DecimalFormat format = new DecimalFormat(pattern);
        return format.format(value);
    }

    private List<String> normalizeHeaders(List<String> rawHeaders, int totalColumns) {
        List<String> paddedHeaders = padRow(rawHeaders, totalColumns);
        Map<String, Integer> duplicateTracker = new LinkedHashMap<>();
        List<String> normalizedHeaders = new ArrayList<>();

        for (int columnIndex = 0; columnIndex < paddedHeaders.size(); columnIndex++) {
            String candidateHeader = paddedHeaders.get(columnIndex);
            String baseHeader = StringUtils.hasText(candidateHeader)
                    ? candidateHeader
                    : "Column " + (columnIndex + 1);
            int duplicateCount = duplicateTracker.getOrDefault(baseHeader, 0) + 1;
            duplicateTracker.put(baseHeader, duplicateCount);
            normalizedHeaders.add(duplicateCount == 1 ? baseHeader : baseHeader + " " + duplicateCount);
        }

        return normalizedHeaders;
    }

    private List<String> createGenericHeaders(int totalColumns) {
        List<String> headers = new ArrayList<>();
        for (int columnIndex = 0; columnIndex < totalColumns; columnIndex++) {
            headers.add("Column " + (columnIndex + 1));
        }
        return headers;
    }

    private List<String> padRow(List<String> rawRow, int totalColumns) {
        List<String> paddedRow = new ArrayList<>(rawRow);
        while (paddedRow.size() < totalColumns) {
            paddedRow.add("");
        }
        return List.copyOf(paddedRow);
    }

    private String describeDelimiter(String delimiter) {
        return switch (delimiter) {
            case "\t" -> "Tab separated";
            case "|" -> "Pipe separated";
            case ";" -> "Semicolon separated";
            case "," -> "Comma separated";
            default -> "Custom separated";
        };
    }

    private void createSummarySheet(Sheet sheet, ParsedTxtData parsedTxtData, Styles styles) {
        createTitleRow(sheet, 2, styles.title(), "Conversion Summary");
        createMetadataRow(sheet, 1, "Source file", parsedTxtData.sourceFileName(), styles);
        createMetadataRow(sheet, 2, "Detected layout", parsedTxtData.delimiterLabel(), styles);
        createMetadataRow(sheet, 3, "Exported rows", String.valueOf(parsedTxtData.rows().size()), styles);
        createMetadataRow(sheet, 4, "Generated at", LocalDateTime.now().format(EXPORT_TIME_FORMAT), styles);
        createMetadataRow(sheet, 5, "Review sheet", DATA_SHEET_NAME, styles);
        if (!parsedTxtData.summaryDetails().isEmpty()) {
            int rowIndex = 7;
            for (Map.Entry<String, String> entry : parsedTxtData.summaryDetails().entrySet()) {
                createMetadataRow(sheet, rowIndex, entry.getKey(), entry.getValue(), styles);
                rowIndex++;
            }
        }

        sheet.setColumnWidth(0, 5_000);
        sheet.setColumnWidth(1, 24_000);
    }

    private int createDataSheetIntro(Sheet sheet, ParsedTxtData parsedTxtData, Styles styles) {
        int nextRowIndex = 0;

        if (StringUtils.hasText(parsedTxtData.dataSheetTitle())) {
            createTitleRow(sheet, parsedTxtData.headers().size(), styles.title(), parsedTxtData.dataSheetTitle());
            nextRowIndex++;
        }

        if (StringUtils.hasText(parsedTxtData.dataSheetSubtitle())) {
            createMergedTextRow(
                    sheet,
                    nextRowIndex,
                    parsedTxtData.headers().size(),
                    styles.subtitle(),
                    parsedTxtData.dataSheetSubtitle());
            nextRowIndex++;
        }

        return nextRowIndex;
    }

    private void createTitleRow(Sheet sheet, int totalColumns, CellStyle titleStyle, String title) {
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(30);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(titleStyle);

        int lastColumnIndex = Math.max(totalColumns - 1, 0);
        if (lastColumnIndex > 0) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastColumnIndex));
        }
    }

    private void createMergedTextRow(Sheet sheet, int rowIndex, int totalColumns, CellStyle style, String text) {
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(22);

        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);

        int lastColumnIndex = Math.max(totalColumns - 1, 0);
        if (lastColumnIndex > 0) {
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, lastColumnIndex));
        }
    }

    private void createMetadataRow(Sheet sheet, int rowIndex, String label, String value, Styles styles) {
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(22);

        Cell labelCell = row.createCell(0);
        labelCell.setCellValue(label);
        labelCell.setCellStyle(styles.metaLabel());

        Cell valueCell = row.createCell(1);
        valueCell.setCellValue(value);
        valueCell.setCellStyle(styles.metaValue());
    }

    private void createHeaderRow(Sheet sheet, int rowIndex, List<String> headers, CellStyle headerStyle) {
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(24);
        for (int columnIndex = 0; columnIndex < headers.size(); columnIndex++) {
            Cell cell = row.createCell(columnIndex);
            cell.setCellValue(headers.get(columnIndex));
            cell.setCellStyle(headerStyle);
        }
    }

    private void createDataRows(
            Sheet sheet,
            int startRowIndex,
            List<List<String>> rows,
            int totalColumns,
            Set<Integer> textOnlyColumns,
            Styles styles) {
        for (int rowOffset = 0; rowOffset < rows.size(); rowOffset++) {
            Row row = sheet.createRow(startRowIndex + rowOffset);
            boolean useStripedStyle = rowOffset % 2 != 0;
            List<String> sourceRow = rows.get(rowOffset);

            for (int columnIndex = 0; columnIndex < totalColumns; columnIndex++) {
                String value = columnIndex < sourceRow.size() ? sourceRow.get(columnIndex) : "";
                writeDataCell(row, columnIndex, value, useStripedStyle, textOnlyColumns.contains(columnIndex), styles);
            }
        }
    }

    private void sizeDataColumns(Sheet sheet, List<String> headers) {
        for (int columnIndex = 0; columnIndex < headers.size(); columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
            int headerWidth = Math.max((headers.get(columnIndex).length() + 4) * 256, 3_200);
            int contentWidth = sheet.getColumnWidth(columnIndex) + 600;
            int preferredWidth = Math.max(headerWidth, contentWidth);
            sheet.setColumnWidth(columnIndex, Math.min(preferredWidth, maxColumnWidth(headers.get(columnIndex))));
        }
    }

    private int maxColumnWidth(String header) {
        return switch (header) {
            case "Terminal/Name", "Fee Jurisdiction", "Routing", "Unparsed Detail" -> 18_000;
            default -> 12_000;
        };
    }

    private void writeDataCell(
            Row row,
            int columnIndex,
            String value,
            boolean striped,
            boolean forceText,
            Styles styles) {
        Cell cell = row.createCell(columnIndex);
        String trimmedValue = value == null ? "" : value.trim();

        if (trimmedValue.isEmpty()) {
            cell.setBlank();
            cell.setCellStyle(striped ? styles.textStriped() : styles.text());
            return;
        }

        BigDecimal numericValue = forceText ? null : parseNumericValue(trimmedValue);
        if (numericValue != null) {
            cell.setCellValue(numericValue.doubleValue());
            cell.setCellStyle(striped ? styles.numberStriped() : styles.number());
            return;
        }

        if ("true".equalsIgnoreCase(trimmedValue) || "false".equalsIgnoreCase(trimmedValue)) {
            cell.setCellValue(Boolean.parseBoolean(trimmedValue));
            cell.setCellStyle(striped ? styles.textStriped() : styles.text());
            return;
        }

        cell.setCellValue(trimmedValue);
        cell.setCellStyle(striped ? styles.textStriped() : styles.text());
    }

    private BigDecimal parseNumericValue(String value) {
        if (!NUMERIC_VALUE_PATTERN.matcher(value).matches()) {
            return null;
        }

        String normalizedValue = value.replace(",", "");
        String unsignedValue = normalizedValue.startsWith("-") ? normalizedValue.substring(1) : normalizedValue;
        boolean hasDecimal = unsignedValue.contains(".");
        String wholeNumberPart = hasDecimal ? unsignedValue.substring(0, unsignedValue.indexOf('.')) : unsignedValue;

        if (!hasDecimal && wholeNumberPart.length() > 10) {
            return null;
        }

        if (wholeNumberPart.length() > 1 && wholeNumberPart.startsWith("0")) {
            return null;
        }

        return new BigDecimal(normalizedValue);
    }

    private BigDecimal parseSignedNumericValue(String value) {
        if (!StringUtils.hasText(value)) {
            return null;
        }

        String normalizedValue = value.trim().replace(",", "");
        boolean isCredit = normalizedValue.endsWith("CR");
        boolean isDebit = normalizedValue.endsWith("DR");
        if (isCredit || isDebit) {
            normalizedValue = normalizedValue.substring(0, normalizedValue.length() - 2);
        }

        if (!NUMERIC_VALUE_PATTERN.matcher(normalizedValue).matches()) {
            return null;
        }

        BigDecimal numericValue = new BigDecimal(normalizedValue);
        return isDebit ? numericValue.negate() : numericValue;
    }

    private Styles createStyles(Workbook workbook) {
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 14);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);

        Font labelFont = workbook.createFont();
        labelFont.setBold(true);

        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(titleStyle);

        CellStyle metaLabelStyle = workbook.createCellStyle();
        metaLabelStyle.setFont(labelFont);
        metaLabelStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        metaLabelStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        applyThinBorders(metaLabelStyle);

        CellStyle metaValueStyle = workbook.createCellStyle();
        metaValueStyle.setAlignment(HorizontalAlignment.LEFT);
        metaValueStyle.setWrapText(true);
        metaValueStyle.setVerticalAlignment(VerticalAlignment.TOP);
        applyThinBorders(metaValueStyle);

        CellStyle subtitleStyle = workbook.createCellStyle();
        subtitleStyle.cloneStyleFrom(metaValueStyle);
        subtitleStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        applyThinBorders(headerStyle);

        CellStyle textStyle = workbook.createCellStyle();
        textStyle.setAlignment(HorizontalAlignment.LEFT);
        textStyle.setWrapText(true);
        textStyle.setVerticalAlignment(VerticalAlignment.TOP);
        applyThinBorders(textStyle);

        CellStyle textStripedStyle = workbook.createCellStyle();
        textStripedStyle.cloneStyleFrom(textStyle);

        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.cloneStyleFrom(textStyle);
        numberStyle.setAlignment(HorizontalAlignment.RIGHT);
        numberStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));

        CellStyle numberStripedStyle = workbook.createCellStyle();
        numberStripedStyle.cloneStyleFrom(numberStyle);

        return new Styles(
                titleStyle,
                metaLabelStyle,
                metaValueStyle,
                subtitleStyle,
                headerStyle,
                textStyle,
                textStripedStyle,
                numberStyle,
                numberStripedStyle);
    }

    private void applyThinBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
    }

    private record ParsedTxtData(
            List<String> headers,
            List<List<String>> rows,
            String delimiterLabel,
            String sourceFileName,
            Set<Integer> textOnlyColumns,
            int frozenColumnCount,
            Map<String, String> summaryDetails,
            String dataSheetTitle,
            String dataSheetSubtitle) {
    }

    private record Styles(
            CellStyle title,
            CellStyle metaLabel,
            CellStyle metaValue,
            CellStyle subtitle,
            CellStyle header,
            CellStyle text,
            CellStyle textStriped,
            CellStyle number,
            CellStyle numberStriped) {
    }

    private record LabelMatch(int start, int end) {
    }
}
