package com.quynq.app.utils;

import java.io.FileOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelExporter {
    private XSSFWorkbook workbook;
    private String outFilename;
    private List<List<Map<String, String>>> sheets;
    private List<String> sheetNames;

    private OutputStream stream;


    public ExcelExporter() {
        workbook = new XSSFWorkbook();
        sheets = new ArrayList<>();
        sheetNames = new ArrayList<>();
    }

    public void setOutputStream(OutputStream stream) {
        this.stream = stream;
    }

    public void setOutputFile(String filename) {
        this.outFilename = filename;
    }

    public void addSheet(List<Map<String, String>> sheetData, String name) {
        sheets.add(sheetData);
        sheetNames.add(name);
    }

    public void export() {
        try {
            if (stream == null) {
                File outputFile = new File(outFilename);
                stream = new FileOutputStream(outputFile);
            }

            int i = 0;
            for (List<Map<String, String>> sheet: sheets) {
                SheetMaker sheetMaker = new SheetMaker(sheet, sheetNames.get(i++), workbook);
                sheetMaker.writeHeaderLine();
                sheetMaker.writeDataLines();
            }

            workbook.write(stream);
            workbook.close();
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
