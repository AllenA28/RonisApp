package dataimporter.implementation.service;

import dataimporter.implementation.model.CellData;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import dataimporter.implementation.utils.DataImporterUtils;
import dataimporter.implementation.utils.ExtendedXSSFSheetXMLHandler;
import dataimporter.implementation.utils.XLSXHeaderFoundException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class ExcelXLSXDataReader {

    private static final DataFormatter formatter = new DataFormatter();

    private ExcelXLSXDataReader() {
    }

    public static List<CellData> readHeaderRow(File excelFile, String sheetName, int headerRowNo) throws DataImporterRuntimeException {
        List<CellData> headerRowData = new ArrayList<>();
        parseExcelFile(excelFile, sheetName, headerRowNo, headerRowData, 0, null);
        return headerRowData;
    }

    public static List<List<CellData>> readDataRow(File excelFile, String sheetName, int startRowIndex, List<CellData> headerRowData) throws DataImporterRuntimeException {
        List<List<CellData>> dataRow = new ArrayList<>();
        parseExcelFile(excelFile, sheetName, 0, headerRowData, startRowIndex, dataRow);
        return dataRow;
    }

    private static void parseExcelFile(File excelFile, String sheetName, int headerRowIndex, List<CellData> headerRowData,
                                       int startRowIndex, List<List<CellData>> dataRow) throws DataImporterRuntimeException {
        final var selectedSheetIndex = new int[]{-1};
        try (XSSFWorkbook workbook = new XSSFWorkbook(excelFile) {
            @Override
            public void parseSheet(java.util.Map<String, XSSFSheet> shIdMap, CTSheet ctSheet) {
                // skipping parsing of any sheet and getting required sheet index
                if (ctSheet.getName().equals(sheetName)) {
                    List<String> keys = new ArrayList<>(shIdMap.keySet());
                    for (int i = 0; i < keys.size(); i++) {
                        Object sheetRId = keys.get(i);
                        if (ctSheet.getId().equals(sheetRId)) {
                            selectedSheetIndex[0] = i;
                        }
                    }
                }
            }
        }) {
            try (var opcPackage = workbook.getPackage()) {
                if (selectedSheetIndex[0] == -1) {
                    DataImporterUtils.logNode.error("Sheet with a name '" + sheetName + "' not found.");
                    throw new DataImporterRuntimeException("Sheet with a name '" + sheetName + "' not found.");
                }
                var strings = new ReadOnlySharedStringsTable(opcPackage, false);
                ContentHandler handler;
                if (!headerRowData.isEmpty()) {
                    handler = new ExtendedXSSFSheetXMLHandler(workbook.getStylesSource(), strings,
                            createSheetHandlerForData(startRowIndex, dataRow, headerRowData, sheetName, workbook.isDate1904()),
                            formatter, false);
                } else {
                    handler = new ExtendedXSSFSheetXMLHandler(workbook.getStylesSource(), strings,
                            createSheetHandlerForHeader(headerRowIndex, headerRowData), formatter, false);
                }
                var sheetParser = XMLHelper.newXMLReader();
                sheetParser.setContentHandler(handler);
                ArrayList<PackagePart> sheets = opcPackage.getPartsByContentType(XSSFRelation.WORKSHEET.getContentType());
                try (var sheet = sheets.get(selectedSheetIndex[0]).getInputStream()) {
                    var sheetSource = new InputSource(sheet);
                    sheetParser.parse(sheetSource);
                }
            }
        } catch (XLSXHeaderFoundException e) {
            // safe to ignore this exception
        } catch (SAXException | ParserConfigurationException | IOException | InvalidFormatException e) {
            throw new DataImporterRuntimeException("Error while opening workbook:", e);
        }
    }

    private static ExtendedXSSFSheetXMLHandler.SheetContentsHandler createSheetHandlerForHeader(int headerRowIndex, List<CellData> headerRowData) {
        return new ExtendedXSSFSheetXMLHandler.SheetContentsHandler() {
            @Override
            public void startRow(int rowNum) {
                headerRowData.clear();
            }

            @Override
            public void endRow(int rowNum) throws XLSXHeaderFoundException {
                if (rowNum == headerRowIndex && !headerRowData.isEmpty() && !headerRowData.stream().allMatch(Objects::isNull)) {
                    throw new XLSXHeaderFoundException("header row #" + (rowNum + 1));
                } else if (rowNum > headerRowIndex) {
                    throw new DataImporterRuntimeException("Unable to find header row!!");
                }
            }

            @Override
            public void cell(String cellReference, String formattedValue, String rawValue, CellType cellType, String formatString, XSSFComment comment, short dataFormat) {
                var cellAddr = new CellAddress(cellReference);
                if (isHeaderRow(cellAddr)) {
                    addToHeaderRow(cellAddr, cellType, formattedValue, cellReference);
                }
            }

            private boolean isHeaderRow(CellAddress cellAddr) {
                return cellAddr.getRow() == headerRowIndex;
            }

            private void addToHeaderRow(CellAddress cellAddr, CellType cellType, String formattedValue, String cellReference) {
                int columnIndex = cellAddr.getColumn();
                switch (cellType) {
                    case FORMULA:
                    case STRING:
                        headerRowData.add(new CellData(columnIndex, cellReference, null, formattedValue.trim()));
                        break;
                    default:
                        headerRowData.add(null);
                }
            }
        };
    }

    private static ExtendedXSSFSheetXMLHandler.SheetContentsHandler createSheetHandlerForData(int startRowIndex, List<List<CellData>> dataRow, List<CellData> headerRowData, String sheetName, boolean isDate1904) {
        return new ExtendedXSSFSheetXMLHandler.SheetContentsHandler() {
            final List<CellData> data = new ArrayList<>();
            boolean isNewRowStarted = false;

            int rowNo;

            @Override
            public void startRow(int rowNum) {
                this.rowNo = rowNum;
                isNewRowStarted = true;
                data.clear();
            }

            @Override
            public void endRow(int rowNum) {
                if (!data.isEmpty()) {
                    List<CellData> cellDataList = new ArrayList<>();
                    for (CellData cellData : data) {
                        cellDataList.add(new CellData(cellData));
                    }
                    dataRow.add(cellDataList);
                    data.clear();
                }
                isNewRowStarted = false;
            }

            @Override
            public void cell(String cellReference, String formattedValue, String rawValue, CellType cellType, String formatString, XSSFComment comment, short dataFormat) {
                try {
                    var cellAddr = new CellAddress(cellReference);
                    int columnIndex = cellAddr.getColumn();
                    final String columnHeader = DataImporterUtils.getColumnHeader(null, headerRowData, columnIndex);

                    if (cellAddr.getRow() >= startRowIndex) {
                        logTrace(cellReference, rawValue, cellType);
                        switch (cellType) {
                            case BOOLEAN:
                                handleBooleanCell(columnIndex, columnHeader, rawValue, formattedValue);
                                break;
                            case ERROR:
                                handleErrorCell(cellAddr, rawValue, columnIndex, columnHeader);
                                break;
                            case FORMULA:
                                handleFormulaCell(columnIndex, columnHeader, rawValue);
                                break;
                            case STRING:
                                handleStringCell(columnIndex, columnHeader, rawValue, formattedValue);
                                break;
                            case NUMERIC:
                                handleNumericCell(cellReference, columnIndex, columnHeader, rawValue, formattedValue, formatString, dataFormat);
                                break;
                            default:
                                data.add(null);
                        }
                    }
                } catch (Exception e) {
                    throw new DataImporterRuntimeException("Unable to read Excel row " + startRowIndex + 1 + " from @Sheet " + sheetName, e);
                }
            }

            private void logTrace(String cellReference, String rawValue, CellType cellType) {
                if (DataImporterUtils.logNode.isTraceEnabled()) {
                    DataImporterUtils.logNode.trace(String.format("Reading %s / '%s' / %s", cellReference, rawValue, cellType.toString()));
                }
            }

            private void handleBooleanCell(int columnIndex, String columnHeader, String rawValue, String formattedValue) {
                data.add(new CellData(columnIndex, columnHeader, rawValue, formattedValue));
            }

            private void handleErrorCell(CellAddress cellAddr, String rawValue, int columnIndex, String columnHeader) {
                if (rawValue.startsWith("#")) { // Check if the error is due to a formula
                    DataImporterUtils.logNode.error("Unable to import data due to invalid formula at cell address " + cellAddr.toString());
                    throw new DataImporterRuntimeException("Unable to import data due to invalid formula at Excel row #" + rowNo);
                }
                data.add(new CellData(columnIndex, columnHeader, rawValue, "ERROR:" + rawValue));
            }

            private void handleFormulaCell(int columnIndex, String columnHeader, String rawValue) {
                data.add(new CellData(columnIndex, columnHeader, rawValue, rawValue));
            }

            private void handleStringCell(int columnIndex, String columnHeader, String rawValue, String formattedValue) {
                data.add(new CellData(columnIndex, columnHeader, rawValue, formattedValue));
            }

            private void handleNumericCell(String cellReference, int columnIndex, String columnHeader, String rawValue, String formattedValue, String formatString, short dataFormat) {
                final var dblCellValue = Double.parseDouble(rawValue);
                if (DataImporterUtils.isCellDateFormatted(dataFormat, formatString) && DateUtil.isValidExcelDate(dblCellValue)) {
                    var dateValue = DateUtil.getJavaDate(dblCellValue, isDate1904);
                    logTraceFormatting(cellReference, rawValue, formatString, formattedValue);
                    data.add(new CellData(columnIndex, columnHeader, dblCellValue, dateValue, formatString, dataFormat));
                } else {
                    data.add(new CellData(columnIndex, columnHeader, rawValue, rawValue));
                }
            }

            private void logTraceFormatting(String cellReference, String rawValue, String formatString, String formattedValue) {
                if (DataImporterUtils.logNode.isTraceEnabled()) {
                    DataImporterUtils.logNode.trace(String.format("Formatting %s / '%s' using format: '%s' as %s", cellReference, rawValue, formatString, formattedValue));
                }
            }
        };
    }
}
