package com.example.excel;

import org.dhatim.fastexcel.writer.*;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.Color;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ExcelService {

    /**
     * Main method to generate Excel file.
     *
     * @param pkpPdf           single PKP info object
     * @param pksPdfList       list of PKS PDF objects
     * @param pkpResultsList   PKP_details
     * @param pksDetailsList   PKS_details (KrfResult)
     * @param carResultsList   Car_results
     * @param excludedCarsList Excluded_cars
     * @param carThresholdsList Cars_thresholds
     * @param response         HTTP response
     * @throws IOException
     */
    public void generateExcel(
            PkpPdf pkpPdf,
            List<PksPdf> pksPdfList,
            List<PkpResults> pkpResultsList,
            List<KrfResult> pksDetailsList,
            List<CarResults> carResultsList,
            List<ExcludedCars> excludedCarsList,
            List<CarThresholds> carThresholdsList,
            HttpServletResponse response
    ) throws IOException {

        // Construct file name based on pkp_name + pkp_date
        String dateString = (pkpPdf.getPkp_date() != null)
                ? pkpPdf.getPkp_date().format(DateTimeFormatter.ISO_DATE) 
                : "unknown_date";
        String fileName = pkpPdf.getPkp_name() + "_" + dateString + ".xlsx";

        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");

        try (ServletOutputStream outputStream = response.getOutputStream();
             Workbook workbook = new Workbook(outputStream, "PKP", "1.0")) {

            // 1) Create Worksheet PKP
            Worksheet wsPKP = workbook.newWorksheet("PKP");
            createPkpSheet(wsPKP, workbook, pkpPdf, pksPdfList);

            // 2) Create Worksheet PKP_details
            Worksheet wsPkpDetails = workbook.newWorksheet("PKP_details");
            createPkpDetailsSheet(wsPkpDetails, workbook, pkpResultsList);

            // 3) Create Worksheet PKS_details
            Worksheet wsPksDetails = workbook.newWorksheet("PKS_details");
            createPksDetailsSheet(wsPksDetails, workbook, pksDetailsList);

            // 4) Create Worksheet Car_results
            Worksheet wsCarResults = workbook.newWorksheet("Car_results");
            createCarResultsSheet(wsCarResults, workbook, carResultsList);

            // 5) Create Worksheet Excluded_cars
            Worksheet wsExcluded = workbook.newWorksheet("Excluded_cars");
            createExcludedCarsSheet(wsExcluded, workbook, excludedCarsList);

            // 6) Create Worksheet Cars_thresholds
            Worksheet wsCarThresholds = workbook.newWorksheet("Cars_thresholds");
            createCarThresholdsSheet(wsCarThresholds, workbook, carThresholdsList);

            // Finish and flush
            workbook.finish();
            outputStream.flush();
        }
    }

    /**
     * Worksheet 1: PKP
     */
    private void createPkpSheet(Worksheet sheet, Workbook workbook, PkpPdf pkpPdf, List<PksPdf> pksPdfList) {

        // Some styles
        Style bold14 = workbook.addStyle().fontSize(14).bold();
        Style bold = workbook.addStyle().bold();
        Style wrapStyle = workbook.addStyle().wrapText(true);
        Style boldWrap = workbook.addStyle().bold().wrapText(true);

        // You could also define these color-coded styles upfront if you prefer
        // For example:
        // Style redFont = workbook.addStyle().fontColor(Color.RED);

        // Row tracker
        int rowIndex = 0;

        // 1) First row: bold 14, pkp_name
        sheet.value(rowIndex, 0, pkpPdf.getPkp_name()).style(bold14);
        // width ~500 for the first column
        sheet.width(0, 500);
        rowIndex++;

        // 2) Second row: pkp_date
        String dateString = (pkpPdf.getPkp_date() != null)
                ? pkpPdf.getPkp_date().format(DateTimeFormatter.ISO_DATE)
                : "";
        sheet.value(rowIndex, 0, dateString);
        rowIndex++;

        // 3) Two empty rows
        rowIndex += 2;

        // 4) Header row for Accuracy, Completeness, Consistency, Timeliness
        //    first cell empty, then 4 bold headers
        //    set each of these columns to width ~150
        sheet.value(rowIndex, 0, "").style(bold);
        sheet.value(rowIndex, 1, "Accuracy").style(boldWrap);
        sheet.value(rowIndex, 2, "Completeness").style(boldWrap);
        sheet.value(rowIndex, 3, "Consistency").style(boldWrap);
        sheet.value(rowIndex, 4, "Timeliness").style(boldWrap);

        sheet.width(1, 150);
        sheet.width(2, 150);
        sheet.width(3, 150);
        sheet.width(4, 150);

        rowIndex++;

        // 5) Row: "System Based" 
        sheet.value(rowIndex, 0, "System Based").style(bold);
        // color-coded cells for pkpPdf accuracy, completeness, etc
        applyColorCodedValue(sheet, workbook, rowIndex, 1, pkpPdf.getAccuracy());
        applyColorCodedValue(sheet, workbook, rowIndex, 2, pkpPdf.getCompleteness());
        applyColorCodedValue(sheet, workbook, rowIndex, 3, pkpPdf.getConsistency());
        applyColorCodedValue(sheet, workbook, rowIndex, 4, pkpPdf.getTimeliness());
        rowIndex++;

        // 6) Row: "Adjusted"
        sheet.value(rowIndex, 0, "Adjusted").style(bold);
        applyColorCodedValue(sheet, workbook, rowIndex, 1, pkpPdf.getAccuracy_amended());
        applyColorCodedValue(sheet, workbook, rowIndex, 2, pkpPdf.getCompleteness_amended());
        applyColorCodedValue(sheet, workbook, rowIndex, 3, pkpPdf.getConsistency_amended());
        applyColorCodedValue(sheet, workbook, rowIndex, 4, pkpPdf.getTimeliness_amended());
        rowIndex++;

        // 7) Two empty rows
        rowIndex += 2;

        // 8) Row: "samochod owner comment:"
        sheet.value(rowIndex, 0, "samochod owner comment:").style(bold);
        rowIndex++;

        // 9) pkp_comment
        sheet.value(rowIndex, 0, pkpPdf.getPkp_comment()).style(wrapStyle);
        rowIndex++;

        // 10) Two empty rows
        rowIndex += 2;

        // 11) Row: "Polska Klasa Samochodow", bold in first cell, 
        //     next columns bold headers "Accuracy", "Completeness", etc.
        sheet.value(rowIndex, 0, "Polska Klasa Samochodow").style(bold);
        sheet.value(rowIndex, 1, "Accuracy").style(boldWrap);
        sheet.value(rowIndex, 2, "Completeness").style(boldWrap);
        sheet.value(rowIndex, 3, "Consistency").style(boldWrap);
        sheet.value(rowIndex, 4, "Timeliness").style(boldWrap);
        rowIndex++;

        // 12) For each PksPdf, create a row
        if (pksPdfList != null) {
            for (PksPdf pks : pksPdfList) {
                // Column 0: pks_name in bold
                sheet.value(rowIndex, 0, pks.getPks_name()).style(bold);
                // Then color-coded cells
                applyColorCodedValue(sheet, workbook, rowIndex, 1, pks.getAccuracy());
                applyColorCodedValue(sheet, workbook, rowIndex, 2, pks.getCompleteness());
                applyColorCodedValue(sheet, workbook, rowIndex, 3, pks.getConsistency());
                applyColorCodedValue(sheet, workbook, rowIndex, 4, pks.getTimeliness());
                rowIndex++;
            }
        }

        // 13) Two empty rows
        rowIndex += 2;

        // 14) Row: "Created at " + pkp_comment_timestamp
        sheet.value(rowIndex, 0, "Created at " + safeString(pkpPdf.getPkp_comment_timestamp()));
        rowIndex++;

        // 15) Row: "uuid " + pkp_comment_uuid
        sheet.value(rowIndex, 0, "uuid " + safeString(pkpPdf.getPkp_comment_uuid()));
        rowIndex++;
    }

    /**
     * Helper to apply color-coded text: red, amber, green, default black
     */
    private void applyColorCodedValue(Worksheet sheet, Workbook workbook, int row, int col, String value) {
        if (value == null) {
            sheet.value(row, col, "");
            return;
        }
        Style style = colorStyle(workbook, value);
        sheet.value(row, col, value).style(style);
    }

    /**
     * Creates a style with the font color depending on 'red', 'amber', 'green' (case-insensitive).
     * Otherwise, black font color.
     */
    private Style colorStyle(Workbook workbook, String ragValue) {
        String val = ragValue.trim().toLowerCase();
        Style style = workbook.addStyle().wrapText(true);

        switch (val) {
            case "red":
                style = style.fontColor(Color.RED);
                break;
            case "amber":
                // a typical amber color: #FFC000 in hex
                style = style.fontColor(new Color(255, 192, 0));
                break;
            case "green":
                // a typical green color: #00B050
                style = style.fontColor(new Color(0, 176, 80));
                break;
            default:
                // black or default
                style = style.fontColor(Color.BLACK);
        }
        return style;
    }

    private String safeString(String input) {
        return (input == null) ? "" : input;
    }

    /**
     * Worksheet 2: PKP_details
     *
     * - PKPResults:
     *   int pkp_id;
     *   LocalDate pkp_date;
     *   String pkp_name;
     *   String dimension;
     *   int red;
     *   int amber;
     *   int green;
     *   int na;
     *   String pkp_status;
     *   String pkp_status_amended;
     */
    private void createPkpDetailsSheet(Worksheet sheet, Workbook workbook, List<PkpResults> pkpResultsList) {
        // Bold style
        Style bold = workbook.addStyle().bold();
        // Wrap style
        Style wrap = workbook.addStyle().wrapText(true);

        // Header
        int row = 0;
        sheet.value(row, 0, "PKP ID").style(bold);
        sheet.value(row, 1, "PKP DATE").style(bold);
        sheet.value(row, 2, "PKP NAME").style(bold);
        sheet.value(row, 3, "DIMENSION").style(bold);
        sheet.value(row, 4, "RED").style(bold);
        sheet.value(row, 5, "AMBER").style(bold);
        sheet.value(row, 6, "GREEN").style(bold);
        sheet.value(row, 7, "NA").style(bold);
        sheet.value(row, 8, "PKP STATUS").style(bold);
        sheet.value(row, 9, "PKP STATUS AMENDED").style(bold);

        // fix width ~500 for PKP NAME column
        sheet.width(2, 500);

        row++;

        // Data
        if (pkpResultsList != null) {
            DateTimeFormatter fmt = DateTimeFormatter.ISO_DATE;
            for (PkpResults pr : pkpResultsList) {
                sheet.value(row, 0, pr.getPkp_id());
                // pkp_date
                String dateString = (pr.getPkp_date() != null) ? pr.getPkp_date().format(fmt) : "";
                sheet.value(row, 1, dateString).style(wrap);

                sheet.value(row, 2, safeString(pr.getPkp_name())).style(wrap);
                sheet.value(row, 3, safeString(pr.getDimension())).style(wrap);
                sheet.value(row, 4, pr.getRed());
                sheet.value(row, 5, pr.getAmber());
                sheet.value(row, 6, pr.getGreen());
                sheet.value(row, 7, pr.getNa());
                sheet.value(row, 8, safeString(pr.getPkp_status())).style(wrap);
                sheet.value(row, 9, safeString(pr.getPkp_status_amended())).style(wrap);
                row++;
            }
        }
    }

    /**
     * Worksheet 3: PKS_details
     *
     * - KRFResult:
     *   int pkp_id;
     *   int pks_id;
     *   LocalDate pkp_date;
     *   String pks_name;
     *   String dimension;
     *   int red;
     *   int amber;
     *   int green;
     *   int na;
     *   String rag_status;
     */
    private void createPksDetailsSheet(Worksheet sheet, Workbook workbook, List<KrfResult> pksDetailsList) {
        Style bold = workbook.addStyle().bold();
        Style wrap = workbook.addStyle().wrapText(true);

        int row = 0;
        // Headers
        sheet.value(row, 0, "PKP ID").style(bold);
        sheet.value(row, 1, "PKS ID").style(bold);
        sheet.value(row, 2, "PKP DATE").style(bold);
        sheet.value(row, 3, "PKS NAME").style(bold);
        sheet.value(row, 4, "DIMENSION").style(bold);
        sheet.value(row, 5, "RED").style(bold);
        sheet.value(row, 6, "AMBER").style(bold);
        sheet.value(row, 7, "GREEN").style(bold);
        sheet.value(row, 8, "NA").style(bold);
        sheet.value(row, 9, "RAG STATUS").style(bold);
        row++;

        // Data
        if (pksDetailsList != null) {
            DateTimeFormatter fmt = DateTimeFormatter.ISO_DATE;
            for (KrfResult kr : pksDetailsList) {
                sheet.value(row, 0, kr.getPkp_id());
                sheet.value(row, 1, kr.getPks_id());
                String dateString = (kr.getPkp_date() != null) ? kr.getPkp_date().format(fmt) : "";
                sheet.value(row, 2, dateString).style(wrap);
                sheet.value(row, 3, safeString(kr.getPks_name())).style(wrap);
                sheet.value(row, 4, safeString(kr.getDimension())).style(wrap);
                sheet.value(row, 5, kr.getRed());
                sheet.value(row, 6, kr.getAmber());
                sheet.value(row, 7, kr.getGreen());
                sheet.value(row, 8, kr.getNa());
                sheet.value(row, 9, safeString(kr.getRag_status())).style(wrap);

                row++;
            }
        }
    }

    /**
     * Worksheet 4: Car_results
     *
     * - CarResults:
     *   int car_id;
     *   String car_name;
     *   String dimension;
     *   String red;    // (some data)
     *   String amber;  // ...
     *   Double car_score;
     *   String car_status;
     */
    private void createCarResultsSheet(Worksheet sheet, Workbook workbook, List<CarResults> carResultsList) {
        Style bold = workbook.addStyle().bold();
        Style wrap = workbook.addStyle().wrapText(true);

        int row = 0;
        // Header
        sheet.value(row, 0, "CAR ID").style(bold);
        sheet.value(row, 1, "CAR NAME").style(bold);
        sheet.value(row, 2, "DIMENSION").style(bold);
        sheet.value(row, 3, "RED").style(bold);
        sheet.value(row, 4, "AMBER").style(bold);
        sheet.value(row, 5, "CAR SCORE").style(bold);
        sheet.value(row, 6, "CAR STATUS").style(bold);
        row++;

        // Data
        if (carResultsList != null) {
            for (CarResults cr : carResultsList) {
                sheet.value(row, 0, cr.getCar_id());
                sheet.value(row, 1, safeString(cr.getCar_name())).style(wrap);
                sheet.value(row, 2, safeString(cr.getDimension())).style(wrap);
                sheet.value(row, 3, safeString(cr.getRed())).style(wrap);
                sheet.value(row, 4, safeString(cr.getAmber())).style(wrap);
                // If car_score is not null
                sheet.value(row, 5, (cr.getCar_score() == null) ? "" : cr.getCar_score());
                sheet.value(row, 6, safeString(cr.getCar_status())).style(wrap);
                row++;
            }
        }
    }

    /**
     * Worksheet 5: Excluded_cars
     *
     * - ExcludedCars:
     *   int car_id;
     *   String car_name;
     *   String exclusion_reason;
     */
    private void createExcludedCarsSheet(Worksheet sheet, Workbook workbook, List<ExcludedCars> excludedCarsList) {
        Style bold = workbook.addStyle().bold();
        Style wrap = workbook.addStyle().wrapText(true);

        int row = 0;
        sheet.value(row, 0, "CAR ID").style(bold);
        sheet.value(row, 1, "CAR NAME").style(bold);
        sheet.value(row, 2, "EXCLUSION REASON").style(bold);
        row++;

        if (excludedCarsList != null) {
            for (ExcludedCars ec : excludedCarsList) {
                sheet.value(row, 0, ec.getCar_id());
                sheet.value(row, 1, safeString(ec.getCar_name())).style(wrap);
                sheet.value(row, 2, safeString(ec.getExclusion_reason())).style(wrap);
                row++;
            }
        }
    }

    /**
     * Worksheet 6: Cars_thresholds
     *
     * - CarThresholds:
     *   String car_name;
     *   int accuracy;
     *   int completeness;
     *   int consistency;
     *   int timeliness;
     */
    private void createCarThresholdsSheet(Worksheet sheet, Workbook workbook, List<CarThresholds> carThresholdsList) {
        Style bold = workbook.addStyle().bold();
        Style wrap = workbook.addStyle().wrapText(true);

        int row = 0;
        sheet.value(row, 0, "CAR NAME").style(bold);
        sheet.value(row, 1, "ACCURACY").style(bold);
        sheet.value(row, 2, "COMPLETENESS").style(bold);
        sheet.value(row, 3, "CONSISTENCY").style(bold);
        sheet.value(row, 4, "TIMELINESS").style(bold);
        row++;

        if (carThresholdsList != null) {
            for (CarThresholds ct : carThresholdsList) {
                sheet.value(row, 0, safeString(ct.getCar_name())).style(wrap);
                sheet.value(row, 1, ct.getAccuracy());
                sheet.value(row, 2, ct.getCompleteness());
                sheet.value(row, 3, ct.getConsistency());
                sheet.value(row, 4, ct.getTimeliness());
                row++;
            }
        }
    }
}
