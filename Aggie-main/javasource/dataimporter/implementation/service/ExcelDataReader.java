package dataimporter.implementation.service;

import dataimporter.implementation.model.CellData;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import dataimporter.implementation.utils.DataImporterUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.io.File;
import java.io.IOException;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class ExcelDataReader implements AutoCloseable {

    private static final String SHEET_IS_NULL = "Sheet is null";
    private static final String ROW_NUMBER_NOT_FOUND = "Row number not found";
    private Workbook workbook;
    private Sheet sheet;

    public ExcelDataReader(File excelFile) throws IOException {
        if (excelFile == null || !excelFile.exists()) {
            DataImporterUtils.logNode.error("Excel file not found.");
            throw new DataImporterRuntimeException("Excel file not found.");
        }
        this.workbook = WorkbookFactory.create(excelFile, null, true);
    }

    public int getNumberOfRows() {
        return sheet.getPhysicalNumberOfRows();
    }

    public boolean hasRow(int rowNo) {
        return sheet.getRow(rowNo) != null;
    }

    public List<CellData> readHeaderRow(int headerRowNo) {
        if (this.sheet == null) {
            DataImporterUtils.logNode.error(SHEET_IS_NULL);
            throw new DataImporterRuntimeException(SHEET_IS_NULL);
        }
        if (!hasRow(headerRowNo)) {
            DataImporterUtils.logNode.error(ROW_NUMBER_NOT_FOUND);
            throw new DataImporterRuntimeException(ROW_NUMBER_NOT_FOUND);
        }
        try (Stream<Cell> cellStream = StreamSupport.stream(sheet.getRow(headerRowNo).spliterator(), false)) {
            return cellStream.sequential().map(cell -> {
                // add column
                Object rawData = getValue(cell, cell.getCellType());
                if (rawData != null) {
                    return evaluateCellData(cell, rawData.toString().trim(), null);
                }
                return null;
            }).filter(Objects::nonNull).collect(Collectors.toList());
        }
    }

    public List<CellData> readDataRow(int currentRowIndex, List<CellData> headerRowData) {
        if (sheet == null) {
            DataImporterUtils.logNode.error(SHEET_IS_NULL);
            throw new DataImporterRuntimeException(SHEET_IS_NULL);
        }
        if (currentRowIndex < 0) {
            throw new DataImporterRuntimeException("Invalid data row index.");
        }
        if (!hasRow(currentRowIndex)) {
            return Collections.emptyList();
        }
        var colIndexes = headerRowData.stream().map(CellData::getColumnIndex).collect(Collectors.toList());
        try (Stream<Cell> cellStream = StreamSupport.stream(sheet.getRow(currentRowIndex).spliterator(), false)) {
            return cellStream.sequential().map(cell -> {
                if (DataImporterUtils.logNode.isTraceEnabled()) {
                    DataImporterUtils.logNode.trace("Reading excel cell " + DataImporterUtils.getCellName(cell) + " from row " + currentRowIndex + 1);
                }
                if (!colIndexes.contains(cell.getColumnIndex())) {
                    return null;
                }
                // add column
                Object rawData = getValue(cell, cell.getCellType());
                if (rawData != null) {
                    return evaluateCellData(cell, rawData, headerRowData);
                }
                return null;
            }).filter(Objects::nonNull).collect(Collectors.toList());
        }
    }

    private Object getValue(Cell cell, CellType cellType) {
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue() ? Boolean.TRUE : Boolean.FALSE;
            case FORMULA:
                return (cell.getCachedFormulaResultType() != null) ? getValue(cell, cell.getCachedFormulaResultType()) : cell.getCellFormula();
            case ERROR:
                return cell.getErrorCellValue();
            case BLANK:
            case _NONE:
            default:
                return null;
        }
    }

    private CellData evaluateCellData(Cell cell, Object cellValueString, List<CellData> headerRowData) {
        final int columnIndex = cell.getColumnIndex();
        final String columnHeader = DataImporterUtils.getColumnHeader(cell, headerRowData, columnIndex);
        switch (cell.getCellType()) {
            case ERROR:
                return new CellData(columnIndex, columnHeader, cellValueString, "ERROR:" + cellValueString);
            case BOOLEAN:
                return new CellData(columnIndex, columnHeader, cellValueString, cellValueString);
            case FORMULA:
                if (cell.getCachedFormulaResultType() == CellType.ERROR) {
                    DataImporterUtils.logNode.error("Unable to import data due to invalid formula at cell address " + cell.getAddress());
                    throw new DataImporterRuntimeException("Unable to import data due to invalid formula at Excel row #" + (cell.getRow().getRowNum() + 1));
                }
                return new CellData(columnIndex, columnHeader, cellValueString, cellValueString);
            case STRING: // We haven't seen this yet.
                var rtsi = new XSSFRichTextString(cellValueString.toString());
                return new CellData(columnIndex, columnHeader, cellValueString, rtsi.toString());
            case NUMERIC:
                final var formatString = cell.getCellStyle().getDataFormatString();
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new CellData(columnIndex, columnHeader, cell.getNumericCellValue(), cell.getDateCellValue(), formatString, cell.getCellStyle().getDataFormat());
                } else {
                    return new CellData(columnIndex, columnHeader, cellValueString, cellValueString);
                }
            default:
                return null;
        }
    }

    public void openSheet(String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) {
            DataImporterUtils.logNode.error("'" + sheetName + "' cannot be empty");
            throw new DataImporterRuntimeException("'" + sheetName + "' cannot be empty");
        }
        if (workbook.getSheet(sheetName) == null) {
            DataImporterUtils.logNode.error("Sheet with a name '" + sheetName + "' not found.");
            throw new DataImporterRuntimeException("Sheet with a name '" + sheetName + "' not found.");
        }
        this.sheet = workbook.getSheet(sheetName);
    }

    @Override
    public void close() throws Exception {
        if (workbook != null) {
            workbook.close();
        }
        workbook = null;
    }
}
