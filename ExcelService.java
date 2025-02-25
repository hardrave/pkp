package com.example.excel;

import org.dhatim.fastexcel.writer.Workbook;
import org.dhatim.fastexcel.writer.Worksheet;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.Color;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ExcelService {

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

            // 1) Worksheet "PKP"
            Worksheet wsPKP = workbook.newWorksheet("PKP");
            createPkpSheet(wsPKP, pkpPdf, pksPdfList);

            // 2) Worksheet "PKP_details"
            Worksheet wsPkpDetails = workbook.newWorksheet("PKP_details");
            createPkpDetailsSheet(wsPkpDetails, pkpResultsList);

            // 3) Worksheet "PKS_details"
            Worksheet wsPksDetails = workbook.newWorksheet("PKS_details");
            createPksDetailsSheet(wsPksDetails, pksDetailsList);

            // 4) Worksheet "Car_results"
            Worksheet wsCarResults = workbook.newWorksheet("Car_results");
            createCarResultsSheet(wsCarResults, carResultsList);

            // 5) Worksheet "Excluded_cars"
            Worksheet wsExcluded = workbook.newWorksheet("Excluded_cars");
            createExcludedCarsSheet(wsExcluded, excludedCarsList);

            // 6) Worksheet "Cars_thresholds"
            Worksheet wsCarThresholds = workbook.newWorksheet("Cars_thresholds");
            createCarThresholdsSheet(wsCarThresholds, carThresholdsList);

            // Finish
            workbook.finish();
            outputStream.flush();
        }
    }

    /**
     * =============================
     * 1) WORKSHEET "PKP"
     * =============================
     *
     * Layout:
     *  Row 1 (bold,14): pkp_name (col0 width ~500)
     *  Row 2: pkp_date
     *  +2 empty rows
     *  Next row: [ "", Accuracy, Completeness, Consistency, Timeliness ] => bold
     *  Next row: [ "System Based", color-coded accuracy, completeness, consistency, timeliness ]
     *  Next row: [ "Adjusted", color-coded accuracy_amended, ... ]
     *  +2 empty rows
     *  Row: "samochod owner comment:" (bold)
     *  Row: pkp_comment
     *  +2 empty rows
     *  Row: "Polska Klasa Samochodow", then bold "Accuracy", "Completeness", "Consistency", "Timeliness"
     *  Rows for pksPdfList
     *  +2 empty rows
     *  Row: "Created at " + pkp_comment_timestamp
     *  Row: "uuid " + pkp_comment_uuid
     */
    private void createPkpSheet(Worksheet sheet, PkpPdf pkpPdf, List<PksPdf> pksPdfList) {
        int row = 0;

        // 1) pkp_name in row 0, col 0 (bold, fontSize=14)
        sheet.value(row, 0, safeString(pkpPdf.getPkp_name()));
        sheet.style(row, 0).bold(true).fontSize(14);
        // approximate "width ~500 px" means a large column width; you may need to tweak
        sheet.width(0, 100); // trial and error

        row++;

        // 2) pkp_date in row 1, col 0
        String dateString = formatDate(pkpPdf.getPkp_date());
        sheet.value(row, 0, dateString);

        row++;

        // 3) two empty rows
        row += 2;

        // 4) header row [ "", Accuracy, Completeness, Consistency, Timeliness ] (bold + wrap)
        sheet.value(row, 0, "");
        sheet.style(row, 0).bold(true);

        sheet.value(row, 1, "Accuracy");
        sheet.style(row, 1).bold(true).wrapText(true);

        sheet.value(row, 2, "Completeness");
        sheet.style(row, 2).bold(true).wrapText(true);

        sheet.value(row, 3, "Consistency");
        sheet.style(row, 3).bold(true).wrapText(true);

        sheet.value(row, 4, "Timeliness");
        sheet.style(row, 4).bold(true).wrapText(true);

        // approximate column widths
        sheet.width(1, 30);
        sheet.width(2, 30);
        sheet.width(3, 30);
        sheet.width(4, 30);

        row++;

        // 5) Row: "System Based"
        sheet.value(row, 0, "System Based");
        sheet.style(row, 0).bold(true);

        applyColorCodedValue(sheet, row, 1, pkpPdf.getAccuracy());
        applyColorCodedValue(sheet, row, 2, pkpPdf.getCompleteness());
        applyColorCodedValue(sheet, row, 3, pkpPdf.getConsistency());
        applyColorCodedValue(sheet, row, 4, pkpPdf.getTimeliness());

        row++;

        // 6) Row: "Adjusted"
        sheet.value(row, 0, "Adjusted");
        sheet.style(row, 0).bold(true);

        applyColorCodedValue(sheet, row, 1, pkpPdf.getAccuracy_amended());
        applyColorCodedValue(sheet, row, 2, pkpPdf.getCompleteness_amended());
        applyColorCodedValue(sheet, row, 3, pkpPdf.getConsistency_amended());
        applyColorCodedValue(sheet, row, 4, pkpPdf.getTimeliness_amended());

        row++;

        // 7) two empty rows
        row += 2;

        // 8) Row: "samochod owner comment:" (bold)
        sheet.value(row, 0, "samochod owner comment:");
        sheet.style(row, 0).bold(true);
        row++;

        // 9) Row: pkp_comment
        sheet.value(row, 0, safeString(pkpPdf.getPkp_comment()));
        sheet.style(row, 0).wrapText(true);
        row++;

        // 10) two empty rows
        row += 2;

        // 11) Row: "Polska Klasa Samochodow", then bold headers
        sheet.value(row, 0, "Polska Klasa Samochodow");
        sheet.style(row, 0).bold(true);

        sheet.value(row, 1, "Accuracy");
        sheet.style(row, 1).bold(true).wrapText(true);

        sheet.value(row, 2, "Completeness");
        sheet.style(row, 2).bold(true).wrapText(true);

        sheet.value(row, 3, "Consistency");
        sheet.style(row, 3).bold(true).wrapText(true);

        sheet.value(row, 4, "Timeliness");
        sheet.style(row, 4).bold(true).wrapText(true);

        row++;

        // 12) Rows for pksPdfList
        if (pksPdfList != null) {
            for (PksPdf pks : pksPdfList) {
                sheet.value(row, 0, safeString(pks.getPks_name()));
                sheet.style(row, 0).bold(true);

                applyColorCodedValue(sheet, row, 1, pks.getAccuracy());
                applyColorCodedValue(sheet, row, 2, pks.getCompleteness());
                applyColorCodedValue(sheet, row, 3, pks.getConsistency());
                applyColorCodedValue(sheet, row, 4, pks.getTimeliness());

                row++;
            }
        }

        // 13) two empty rows
        row += 2;

        // 14) "Created at " + pkp_comment_timestamp
        sheet.value(row, 0, "Created at " + safeString(pkpPdf.getPkp_comment_timestamp()));
        row++;

        // 15) "uuid " + pkpPdf.getPkp_comment_uuid()
        sheet.value(row, 0, "uuid " + safeString(pkpPdf.getPkp_comment_uuid()));
        row++;
    }

    /**
     * Helper method to apply color coding for red/amber/green (all others => black).
     */
    private void applyColorCodedValue(Worksheet sheet, int row, int col, String value) {
        if (value == null) {
            sheet.value(row, col, "");
            return;
        }
        sheet.value(row, col, value);
        sheet.style(row, col).wrapText(true); // always wrap if large

        String lower = value.trim().toLowerCase();
        switch (lower) {
            case "red":
                sheet.style(row, col).fontColor(Color.RED);
                break;
            case "amber":
                // typical amber color
                sheet.style(row, col).fontColor(new Color(255, 192, 0));
                break;
            case "green":
                sheet.style(row, col).fontColor(new Color(0, 176, 80));
                break;
            default:
                sheet.style(row, col).fontColor(Color.BLACK);
        }
    }

    /**
     * =============================
     * 2) WORKSHEET "PKP_details"
     * =============================
     *
     * Fields in PkpResults:
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

        // Header row
        sheet.value(row, 0, "PKP ID"); sheet.style(row, 0).bold(true);
        sheet.value(row, 1, "PKP DATE"); sheet.style(row, 1).bold(true);
        sheet.value(row, 2, "PKP NAME"); sheet.style(row, 2).bold(true);
        sheet.value(row, 3, "DIMENSION"); sheet.style(row, 3).bold(true);
        sheet.value(row, 4, "RED"); sheet.style(row, 4).bold(true);
        sheet.value(row, 5, "AMBER"); sheet.style(row, 5).bold(true);
        sheet.value(row, 6, "GREEN"); sheet.style(row, 6).bold(true);
        sheet.value(row, 7, "NA"); sheet.style(row, 7).bold(true);
        sheet.value(row, 8, "PKP STATUS"); sheet.style(row, 8).bold(true);
        sheet.value(row, 9, "PKP STATUS AMENDED"); sheet.style(row, 9).bold(true);

        // Large column width for PKP NAME
        sheet.width(2, 80);

        row++;

        // Data
        if (pkpResultsList != null) {
            for (PkpResults pr : pkpResultsList) {
                sheet.value(row, 0, pr.getPkp_id());

                String dateString = formatDate(pr.getPkp_date());
                sheet.value(row, 1, dateString);
                sheet.style(row, 1).wrapText(true);

                sheet.value(row, 2, safeString(pr.getPkp_name()));
                sheet.style(row, 2).wrapText(true);

                sheet.value(row, 3, safeString(pr.getDimension()));
                sheet.style(row, 3).wrapText(true);

                sheet.value(row, 4, pr.getRed());
                sheet.value(row, 5, pr.getAmber());
                sheet.value(row, 6, pr.getGreen());
                sheet.value(row, 7, pr.getNa());

                sheet.value(row, 8, safeString(pr.getPkp_status()));
                sheet.style(row, 8).wrapText(true);

                sheet.value(row, 9, safeString(pr.getPkp_status_amended()));
                sheet.style(row, 9).wrapText(true);

                row++;
            }
        }
    }

    /**
     * =============================
     * 3) WORKSHEET "PKS_details"
     * =============================
     *
     * KrfResult:
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
        sheet.value(row, 0, "PKP ID").bold(true);
        sheet.value(row, 1, "PKS ID").bold(true);
        sheet.value(row, 2, "PKP DATE").bold(true);
        sheet.value(row, 3, "PKS NAME").bold(true);
        sheet.value(row, 4, "DIMENSION").bold(true);
        sheet.value(row, 5, "RED").bold(true);
        sheet.value(row, 6, "AMBER").bold(true);
        sheet.value(row, 7, "GREEN").bold(true);
        sheet.value(row, 8, "NA").bold(true);
        sheet.value(row, 9, "RAG STATUS").bold(true);

        row++;

        if (pksDetailsList != null) {
            for (KrfResult kr : pksDetailsList) {
                sheet.value(row, 0, kr.getPkp_id());
                sheet.value(row, 1, kr.getPks_id());

                String dateString = formatDate(kr.getPkp_date());
                sheet.value(row, 2, dateString);
                sheet.style(row, 2).wrapText(true);

                sheet.value(row, 3, safeString(kr.getPks_name()));
                sheet.style(row, 3).wrapText(true);

                sheet.value(row, 4, safeString(kr.getDimension()));
                sheet.style(row, 4).wrapText(true);

                sheet.value(row, 5, kr.getRed());
                sheet.value(row, 6, kr.getAmber());
                sheet.value(row, 7, kr.getGreen());
                sheet.value(row, 8, kr.getNa());

                sheet.value(row, 9, safeString(kr.getRag_status()));
                sheet.style(row, 9).wrapText(true);

                row++;
            }
        }
    }

    /**
     * =============================
     * 4) WORKSHEET "Car_results"
     * =============================
     *
     * CarResults:
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
        sheet.value(row, 0, "CAR ID").bold(true);
        sheet.value(row, 1, "CAR NAME").bold(true);
        sheet.value(row, 2, "DIMENSION").bold(true);
        sheet.value(row, 3, "RED").bold(true);
        sheet.value(row, 4, "AMBER").bold(true);
        sheet.value(row, 5, "CAR SCORE").bold(true);
        sheet.value(row, 6, "CAR STATUS").bold(true);

        row++;

        // Data
        if (carResultsList != null) {
            for (CarResults cr : carResultsList) {
                sheet.value(row, 0, cr.getCar_id());
                sheet.value(row, 1, safeString(cr.getCar_name()));
                sheet.style(row, 1).wrapText(true);

                sheet.value(row, 2, safeString(cr.getDimension()));
                sheet.style(row, 2).wrapText(true);

                sheet.value(row, 3, safeString(cr.getRed()));
                sheet.style(row, 3).wrapText(true);

                sheet.value(row, 4, safeString(cr.getAmber()));
                sheet.style(row, 4).wrapText(true);

                sheet.value(row, 5, (cr.getCar_score() == null ? "" : cr.getCar_score()));

                sheet.value(row, 6, safeString(cr.getCar_status()));
                sheet.style(row, 6).wrapText(true);

                row++;
            }
        }
    }

    /**
     * =============================
     * 5) WORKSHEET "Excluded_cars"
     * =============================
     *
     * ExcludedCars:
     *   int car_id;
     *   String car_name;
     *   String exclusion_reason;
     */
    private void createExcludedCarsSheet(Worksheet sheet, List<ExcludedCars> excludedCarsList) {
        int row = 0;

        sheet.value(row, 0, "CAR ID").bold(true);
        sheet.value(row, 1, "CAR NAME").bold(true);
        sheet.value(row, 2, "EXCLUSION REASON").bold(true);
        row++;

        if (excludedCarsList != null) {
            for (ExcludedCars ec : excludedCarsList) {
                sheet.value(row, 0, ec.getCar_id());

                sheet.value(row, 1, safeString(ec.getCar_name()));
                sheet.style(row, 1).wrapText(true);

                sheet.value(row, 2, safeString(ec.getExclusion_reason()));
                sheet.style(row, 2).wrapText(true);

                row++;
            }
        }
    }

    /**
     * =============================
     * 6) WORKSHEET "Cars_thresholds"
     * =============================
     *
     * CarThresholds:
     *   String car_name;
     *   int accuracy;
     *   int completeness;
     *   int consistency;
     *   int timeliness;
     */
    private void createCarThresholdsSheet(Worksheet sheet, List<CarThresholds> carThresholdsList) {
        int row = 0;

        sheet.value(row, 0, "CAR NAME").bold(true);
        sheet.value(row, 1, "ACCURACY").bold(true);
        sheet.value(row, 2, "COMPLETENESS").bold(true);
        sheet.value(row, 3, "CONSISTENCY").bold(true);
        sheet.value(row, 4, "TIMELINESS").bold(true);
        row++;

        if (carThresholdsList != null) {
            for (CarThresholds ct : carThresholdsList) {
                sheet.value(row, 0, safeString(ct.getCar_name()));
                sheet.style(row, 0).wrapText(true);

                sheet.value(row, 1, ct.getAccuracy());
                sheet.value(row, 2, ct.getCompleteness());
                sheet.value(row, 3, ct.getConsistency());
                sheet.value(row, 4, ct.getTimeliness());
                row++;
            }
        }
    }

    // --- Utilities --- //

    private String safeString(String val) {
        return (val == null) ? "" : val;
    }

    private String formatDate(LocalDate date) {
        return (date == null) ? "" : date.format(DateTimeFormatter.ISO_DATE);
    }
}
