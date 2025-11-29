package org.project.app.export;

import jakarta.transaction.Transactional;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.project.app.program.orphan.domain.OrphanApplication;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

@Slf4j
@RequiredArgsConstructor
@Transactional
@Service
public class ExcelExportService {
    public ByteArrayInputStream exportAsStream(List<OrphanApplication> data, List<String> headers) {
        Workbook workbook = exportDynamic(data, headers);

        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.write(out);
            workbook.close();
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            log.error("Excel export failed", e);
            throw new RuntimeException("Excel export failed", e);
        }
    }

    public Workbook exportDynamic(List<?> data, List<String> fields) {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Export");

        // Create bold header style
        CellStyle headerStyle = createHeaderStyle(workbook);

        // HEADER ROW
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < fields.size(); i++) {
            String raw = fields.get(i);
            String pretty = beautifyHeader(raw);

            Cell cell = headerRow.createCell(i);
            cell.setCellValue(pretty);
            cell.setCellStyle(headerStyle);
        }

        // BODY
        for (int rowIndex = 0; rowIndex < data.size(); rowIndex++) {
            Row row = sheet.createRow(rowIndex + 1);
            Object item = data.get(rowIndex);

            for (int col = 0; col < fields.size(); col++) {
                String fieldPath = fields.get(col);
                Object val = readNestedField(item, fieldPath);

                row.createCell(col).setCellValue(
                        val != null ? val.toString() : ""
                );
            }
        }

        // Auto-size columns
        for (int col = 0; col < fields.size(); col++) {
            sheet.autoSizeColumn(col);
        }

        return workbook;
    }

    private Object readNestedField(Object obj, String path) {
        try {
            String[] parts = path.split("\\.");

            Object current = obj;

            for (String part : parts) {
                if (current == null) return null;

                Field field = null;
                Class<?> clazz = current.getClass();

                // walk up inheritance hierarchy
                while (clazz != null) {
                    try {
                        field = clazz.getDeclaredField(part);
                        break;
                    } catch (NoSuchFieldException ex) {
                        clazz = clazz.getSuperclass();
                    }
                }

                if (field == null) return null;

                field.setAccessible(true);
                current = field.get(current);
            }

            return current;
        } catch (Exception e) {
            return null;
        }
    }

    // Helpers
    private String beautifyHeader(String raw) {
        // Take only the last part after dot-notation
        String clean = raw.contains(".")
                ? raw.substring(raw.lastIndexOf(".") + 1)
                : raw;

        // Insert spaces before capitals: fathersName → fathers Name
        clean = clean.replaceAll("([A-Z])", " $1").trim();

        // Capitalize first letter
        return clean.substring(0, 1).toUpperCase() + clean.substring(1);
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }



}
