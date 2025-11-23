package org.project.app.export;

import lombok.Data;

import java.util.List;

@Data
public class ExcelExportRequest {
    public List<String> headers;
}
