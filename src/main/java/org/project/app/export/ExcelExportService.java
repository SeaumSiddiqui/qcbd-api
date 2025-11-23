package org.project.app.export;

import jakarta.transaction.Transactional;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

        // Header
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < fields.size(); i++) {
            String fieldName = fields.get(i);

            // Remove class name (dot-notation)
            String cleanName = fieldName.contains(".")
                    ? fieldName.substring(fieldName.lastIndexOf(".") + 1)
                    : fieldName;

            headerRow.createCell(i).setCellValue(cleanName);
        }

        // Body
        for (int rowIndex = 0; rowIndex < data.size(); rowIndex++) {
            Row row = sheet.createRow(rowIndex + 1);
            Object item = data.get(rowIndex);

            for (int col = 0; col < fields.size(); col++) {
                String fieldName = fields.get(col);
                Object val = readNestedField(item, fieldName);

                row.createCell(col).setCellValue(
                        val != null ? val.toString() : ""
                );
            }
        }

        // Auto-size AFTER all data is placed
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

}
