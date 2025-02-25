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
     * @param pkpPdf            single PKP info object
     * @param pksPdfList        list of PKS PDF objects
     * @param pkpResultsList    PKP_details
     * @param pksDetailsList    PKS_details (KrfResult)
     * @param carResultsList    Car_results
     * @param excludedCarsList  Excluded_cars
     * @param carThresholdsList Cars_thresholds
     * @param response          HTTP response
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
            createPkpSheet(wsPKP, pkpPdf, pksPdfList);

            // 2) Create Worksheet PKP_details
            Worksheet wsPkpDetails = workbook.newWorksheet("PKP_details");
            createPkpDetailsSheet(wsPkpDetails, pkpResultsList);

            // 3) Create Worksheet PKS_details
            Worksheet wsPksDetails = workbook.newWorksheet("PKS_details");
            createPksDetailsSheet(wsPksDetails, pksDetailsList);

            // 4) Create Worksheet Car_results
            Worksheet wsCarResults = workbook.newWorksheet("Car_results");
            createCarResultsSheet(wsCarResults, carResultsList);

            // 5) Create Worksheet Excluded_cars
            Worksheet wsExcluded = workbook.newWorksheet("Excluded_cars");
            createExcludedCarsSheet(wsExcluded, excludedCarsList);

            // 6) Create Worksheet Cars_thresholds
            Worksheet wsCarThresholds = workbook.newWorksheet("Cars_thresholds");
            createCarThresholdsSheet(wsCarThresholds, carThresholdsList);

            // Finish and flush
            workbook.finish();
            outputStream.flush();
        }
    }

    /**
     * Worksheet 1: PKP
     *
     * Layout:
     *  1) Row 1 (bold, font size 14): pkp_name (col width ~500)
     *  2) Row 2: pkp_date
     *  3) Two empty rows
     *  4) Header row: [empty, Accuracy, Completeness, Consistency, Timeliness] (width ~150 each)
     *  5) Row: "System Based", color-coded values
     *  6) Row: "Adjusted", color-coded values
     *  7) Two empty rows
     *  8) Row: "samochod owner comment:"
     *  9) Row: pkp_comment
     *  10) Two empty rows
     *  11) Row: "Polska Klasa Samochodow", then columns [Accuracy, Completeness, Consistency, Timeliness] in bold
     *  12) Each PksPdf in pksPdfList => row with pks_name(bold), color-coded A/C/C/T
     *  13) Two empty rows
     *  14) Row: "Created at " + pkp_comment_timestamp
     *  15) Row: "uuid " + pkp_comment_uuid
     */
    private void createPkpSheet(Worksheet sheet, PkpPdf pkpPdf, List<PksPdf> pksPdfList) {

        // Row tracker
        int row = 0;

        // 1) Row 1: pkp_name in bold, font size 14
        sheet.value(row, 0, safeString(pkpPdf.getPkp_name()))
                .bold()
                .fontSize(14);
        // Approximate column width for the first column (~500 px is not direct, we guess a big enough width):
        sheet.width(0, 100); // Adjust as needed

        row++;

        // 2) Row 2: pkp_date
        String dateString = (pkpPdf.getPkp_date() != null)
                ? pkpPdf.getPkp_date().format(DateTimeFormatter.ISO_DATE)
                : "";
        sheet.value(row, 0, dateString);
        row++;

        // 3) Two empty rows
        row += 2;

        // 4) Headers row: first cell empty, next 4 bold with wrap & ~150 width
        sheet.value(row, 0, "").bold();
        sheet.value(row, 1, "Accuracy").bold().wrapText(true);
        sheet.value(row, 2, "Completeness").bold().wrapText(true);
        sheet.value(row, 3, "Consistency").bold().wrapText(true);
        sheet.value(row, 4, "Timeliness").bold().wrapText(true);

        // Set column widths (approx):
        sheet.width(1, 30);
        sheet.width(2, 30);
        sheet.width(3, 30);
        sheet.width(4, 30);

        row++;

        // 5) Row: "System Based"
        sheet.value(row, 0, "System Based").bold();
        applyColorCodedValue(sheet, row, 1, pkpPdf.getAccuracy());
        applyColorCodedValue(sheet, row, 2, pkpPdf.getCompleteness());
        applyColorCodedValue(sheet, row, 3, pkpPdf.getConsistency());
        applyColorCodedValue(sheet, row, 4, pkpPdf.getTimeliness());
        row++;

        // 6) Row: "Adjusted"
        sheet.value(row, 0, "Adjusted").bold();
        applyColorCodedValue(sheet, row, 1, pkpPdf.getAccuracy_amended());
        applyColorCodedValue(sheet, row, 2, pkpPdf.getCompleteness_amended());
        applyColorCodedValue(sheet, row, 3, pkpPdf.getConsistency_amended());
        applyColorCodedValue(sheet, row, 4, pkpPdf.getTimeliness_amended());
        row++;

        // 7) Two empty rows
        row += 2;

        // 8) Row: "samochod owner comment:"
        sheet.value(row, 0, "samochod owner comment:").bold();
        row++;

        // 9) pkp_comment
        sheet.value(row, 0, safeString(pkpPdf.getPkp_comment()))
                .wrapText(true);
        row++;

        // 10) Two empty rows
        row += 2;

        // 11) Row: "Polska Klasa Samochodow", then bold headers A/C/C/T
        sheet.value(row, 0, "Polska Klasa Samochodow").bold();
        sheet.value(row, 1, "Accuracy").bold().wrapText(true);
        sheet.value(row, 2, "Completeness").bold().wrapText(true);
        sheet.value(row, 3, "Consistency").bold().wrapText(true);
        sheet.value(row, 4, "Timeliness").bold().wrapText(true);
        row++;

        // 12) Rows for each PksPdf
        if (pksPdfList != null) {
            for (PksPdf pks : pksPdfList) {
                sheet.value(row, 0, safeString(pks.getPks_name())).bold();
                applyColorCodedValue(sheet, row, 1, pks.getAccuracy());
                applyColorCodedValue(sheet, row, 2, pks.getCompleteness());
                applyColorCodedValue(sheet, row, 3, pks.getConsistency());
                applyColorCodedValue(sheet, row, 4, pks.getTimeliness());
                row++;
            }
        }

        // 13) Two empty rows
        row += 2;

        // 14) "Created at " + pkp_comment_timestamp
        sheet.value(row, 0, "Created at " + safeString(pkpPdf.getPkp_comment_timestamp()));
        row++;

        // 15) "uuid " + pkp_comment_uuid
        sheet.value(row, 0, "uuid " + safeString(pkpPdf.getPkp_comment_uuid()));
        row++;
    }

    /**
     * Helper for color-coding a single cell:
     *  - red -> .fontColor(Color.RED)
     *  - amber -> .fontColor(new Color(255, 192, 0))
     *  - green -> .fontColor(new Color(0, 176, 80))
     *  - else -> black
     */
    private void applyColorCodedValue(Worksheet sheet, int row, int col, String rawValue) {
        if (rawValue == null) {
            sheet.value(row, col, "");
            return;
        }
        String lower = rawValue.trim().toLowerCase();

        CellData cell = sheet.value(row, col, rawValue)
                .wrapText(true); // always wrap

        switch (lower) {
            case "red":
                cell.fontColor(Color.RED);
                break;
            case "amber":
                cell.fontColor(new Color(255, 192, 0));
                break;
            case "green":
                cell.fontColor(new Color(0, 176, 80));
                break;
            default:
                cell.fontColor(Color.BLACK);
        }
    }

    /**
     * Worksheet 2: PKP_details
     *
     * - PkpResults:
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
    private void createPkpDetailsSheet(Worksheet sheet, List<PkpResults> pkpResultsList) {
        int row = 0;

        // Header row (bold)
        sheet.value(row, 0, "PKP ID").bold();
        sheet.value(row, 1, "PKP DATE").bold();
        sheet.value(row, 2, "PKP NAME").bold();
        sheet.value(row, 3, "DIMENSION").bold();
        sheet.value(row, 4, "RED").bold();
        sheet.value(row, 5, "AMBER").bold();
        sheet.value(row, 6, "GREEN").bold();
        sheet.value(row, 7, "NA").bold();
        sheet.value(row, 8, "PKP STATUS").bold();
        sheet.value(row, 9, "PKP STATUS AMENDED").bold();

        // Approximate large width for PKP NAME
        sheet.width(2, 100); // adjust as needed

        row++;

        // Data rows
        if (pkpResultsList != null) {
            DateTimeFormatter fmt = DateTimeFormatter.ISO_DATE;
            for (PkpResults pr : pkpResultsList) {
                sheet.value(row, 0, pr.getPkp_id());
                String dateString = (pr.getPkp_date() != null) ? pr.getPkp_date().format(fmt) : "";
                sheet.value(row, 1, dateString).wrapText(true);

                sheet.value(row, 2, safeString(pr.getPkp_name())).wrapText(true);
                sheet.value(row, 3, safeString(pr.getDimension())).wrapText(true);
                sheet.value(row, 4, pr.getRed());
                sheet.value(row, 5, pr.getAmber());
                sheet.value(row, 6, pr.getGreen());
                sheet.value(row, 7, pr.getNa());
                sheet.value(row, 8, safeString(pr.getPkp_status())).wrapText(true);
                sheet.value(row, 9, safeString(pr.getPkp_status_amended())).wrapText(true);

                row++;
            }
        }
    }

    /**
     * Worksheet 3: PKS_details
     *
     * - KrfResult:
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
    private void createPksDetailsSheet(Worksheet sheet, List<KrfResult> pksDetailsList) {
        int row = 0;

        // Headers
        sheet.value(row, 0, "PKP ID").bold();
        sheet.value(row, 1, "PKS ID").bold();
        sheet.value(row, 2, "PKP DATE").bold();
        sheet.value(row, 3, "PKS NAME").bold();
        sheet.value(row, 4, "DIMENSION").bold();
        sheet.value(row, 5, "RED").bold();
        sheet.value(row, 6, "AMBER").bold();
        sheet.value(row, 7, "GREEN").bold();
        sheet.value(row, 8, "NA").bold();
        sheet.value(row, 9, "RAG STATUS").bold();
        row++;

        // Data
        if (pksDetailsList != null) {
            DateTimeFormatter fmt = DateTimeFormatter.ISO_DATE;
            for (KrfResult kr : pksDetailsList) {
                sheet.value(row, 0, kr.getPkp_id());
                sheet.value(row, 1, kr.getPks_id());
                String dateString = (kr.getPkp_date() != null) ? kr.getPkp_date().format(fmt) : "";
                sheet.value(row, 2, dateString).wrapText(true);
                sheet.value(row, 3, safeString(kr.getPks_name())).wrapText(true);
                sheet.value(row, 4, safeString(kr.getDimension())).wrapText(true);
                sheet.value(row, 5, kr.getRed());
                sheet.value(row, 6, kr.getAmber());
                sheet.value(row, 7, kr.getGreen());
                sheet.value(row, 8, kr.getNa());
                sheet.value(row, 9, safeString(kr.getRag_status())).wrapText(true);

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
     *   String red;
     *   String amber;
     *   Double car_score;
     *   String car_status;
     */
    private void createCarResultsSheet(Worksheet sheet, List<CarResults> carResultsList) {
        int row = 0;

        // Header
        sheet.value(row, 0, "CAR ID").bold();
        sheet.value(row, 1, "CAR NAME").bold();
        sheet.value(row, 2, "DIMENSION").bold();
        sheet.value(row, 3, "RED").bold();
        sheet.value(row, 4, "AMBER").bold();
        sheet.value(row, 5, "CAR SCORE").bold();
        sheet.value(row, 6, "CAR STATUS").bold();
        row++;

        // Data
        if (carResultsList != null) {
            for (CarResults cr : carResultsList) {
                sheet.value(row, 0, cr.getCar_id());
                sheet.value(row, 1, safeString(cr.getCar_name())).wrapText(true);
                sheet.value(row, 2, safeString(cr.getDimension())).wrapText(true);
                sheet.value(row, 3, safeString(cr.getRed())).wrapText(true);
                sheet.value(row, 4, safeString(cr.getAmber())).wrapText(true);
                sheet.value(row, 5, (cr.getCar_score() != null) ? cr.getCar_score() : "");
                sheet.value(row, 6, safeString(cr.getCar_status())).wrapText(true);
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
    private void createExcludedCarsSheet(Worksheet sheet, List<ExcludedCars> excludedCarsList) {
        int row = 0;

        sheet.value(row, 0, "CAR ID").bold();
        sheet.value(row, 1, "CAR NAME").bold();
        sheet.value(row, 2, "EXCLUSION REASON").bold();
        row++;

        if (excludedCarsList != null) {
            for (ExcludedCars ec : excludedCarsList) {
                sheet.value(row, 0, ec.getCar_id());
                sheet.value(row, 1, safeString(ec.getCar_name())).wrapText(true);
                sheet.value(row, 2, safeString(ec.getExclusion_reason())).wrapText(true);
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
    private void createCarThresholdsSheet(Worksheet sheet, List<CarThresholds> carThresholdsList) {
        int row = 0;

        sheet.value(row, 0, "CAR NAME").bold();
        sheet.value(row, 1, "ACCURACY").bold();
        sheet.value(row, 2, "COMPLETENESS").bold();
        sheet.value(row, 3, "CONSISTENCY").bold();
        sheet.value(row, 4, "TIMELINESS").bold();
        row++;

        if (carThresholdsList != null) {
            for (CarThresholds ct : carThresholdsList) {
                sheet.value(row, 0, safeString(ct.getCar_name())).wrapText(true);
                sheet.value(row, 1, ct.getAccuracy());
                sheet.value(row, 2, ct.getCompleteness());
                sheet.value(row, 3, ct.getConsistency());
                sheet.value(row, 4, ct.getTimeliness());
                row++;
            }
        }
    }

    /**
     * Utility to avoid null pointer on strings.
     */
    private String safeString(String val) {
        return (val == null) ? "" : val;
    }
}
