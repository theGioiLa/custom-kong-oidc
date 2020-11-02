package com.quynq.app.utils;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetMaker {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private String sheetName;
    private List<Map<String, String>> data;

    private List<String> headers;

    public SheetMaker(List<Map<String, String>> data, String sheetName, XSSFWorkbook workbook) {
        this.data = data;
        this.sheetName = sheetName;
        this.workbook = workbook;
        extractHeader();
    }

    void extractHeader() {
        if (!data.isEmpty()) {
            headers = new ArrayList<>();
            for(String header: data.get(0).keySet()) {
                headers.add(header);
            }
        } 
    }

    public void writeHeaderLine() {
        sheet = workbook.createSheet(sheetName);

        sheet.autoSizeColumn(headers.size());
        Row row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);

        fillRow(row, headers, style);
    }

    public void fillRow(Row row, List<String> dataRow, CellStyle style) {
        for (int columnID = 0; columnID < dataRow.size(); columnID++) {
            fillCell(row, columnID, dataRow.get(columnID), style);
        }
    }

    protected void fillCell(Row row, int columnID, Object value, CellStyle style) {
            Cell cell = row.createCell(columnID);
            cell.setCellValue((String) value);
            cell.setCellStyle(style);
    }

    protected void writeDataLines() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);

        int rowId = 1;
        int columnID = 0;
        for (Map<String, String> dataRow: data) {
            Row row = sheet.createRow(rowId);
            columnID = 0;
            for (String value: dataRow.values()) {
                fillCell(row, columnID++, value, style);
            }

            rowId++;
        }
    }
}
