package com.vip.dsa;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import java.util.stream.Stream;

public class CombineSheets {
    private final CellCopyPolicy cellCopyPolicy;
    private final CellCopyContext cellCopyContext;

    public CombineSheets() {
        cellCopyPolicy = getCellCopyPolicy();
        cellCopyContext = new CellCopyContext();
    }

    private CellCopyPolicy getCellCopyPolicy() {
        return new CellCopyPolicy.Builder()
                .cellValue(true)
                .cellStyle(true)
                .cellFormula(true)
                .copyHyperlink(true)
                .mergeHyperlink(true)
                .rowHeight(true)
                .condenseRows(true)
                .mergedRegions(true).build();
    }

    public void mergeAllSheets(String filePath) throws IOException {
        File file = new File(filePath);
        try {
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("All Topics");
            int row = 0;

            FileInputStream inputStream = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(inputStream);
            for (int i = 0; i < workbook.getNumberOfSheets() - 1; i++) {
                Sheet oldSheet = workbook.getSheetAt(i);
                if (row == 0) {
                    addHeader(oldSheet, newSheet);
                    row++;
                }
                for (int j = 1; j <= oldSheet.getLastRowNum(); j++) {
                    Row oldRow = oldSheet.getRow(j);
                    Row newRow = newSheet.createRow(row);
                    newRow.createCell(0).setCellValue(row);
                    newRow.createCell(1).setCellValue(oldSheet.getSheetName());
                    for (int k = 1; k < oldRow.getPhysicalNumberOfCells(); k++) {
                        Cell oldCell = oldRow.getCell(k);
                        Cell newCell = newRow.createCell(k + 1);
                        CellUtil.copyCell(oldCell, newCell, cellCopyPolicy, cellCopyContext);
                    }
                    row++;
                }
            }

            FileOutputStream outputStream = new FileOutputStream(file.getParent() + File.separator + "DSA_Bank_Merged.xlsx");
            newWorkbook.write(outputStream);

            inputStream.close();
            outputStream.close();
        } catch (FileNotFoundException e) {
            System.out.println("File not found at: " + file.getAbsolutePath());
            throw e;
        } catch (IOException e) {
            System.out.println("Unable to read excel file: " + file.getName());
            throw e;
        }
    }

    private void addHeader(Sheet oldSheet, Sheet newSheet) {
        Row oldRow = oldSheet.getRow(0);
        Row newRow = newSheet.createRow(0);

        Cell oldSeqCell = oldRow.getCell(0);
        Cell newSeqCell = newRow.createCell(0);
        CellUtil.copyCell(oldSeqCell, newSeqCell, cellCopyPolicy, cellCopyContext);

        Cell topicCell = newRow.createCell(1);
        CellUtil.copyCell(oldSeqCell, topicCell, cellCopyPolicy, cellCopyContext);
        topicCell.setCellValue("Topic");

        for (int i = 1; i < oldRow.getPhysicalNumberOfCells(); i++) {
            Cell oldCell = oldRow.getCell(i);
            Cell newCell = newRow.createCell(i + 1);
            CellUtil.copyCell(oldCell, newCell, cellCopyPolicy, cellCopyContext);
        }
    }
}
