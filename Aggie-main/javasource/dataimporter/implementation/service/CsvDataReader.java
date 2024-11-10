package dataimporter.implementation.service;

import dataimporter.implementation.model.CellData;
import dataimporter.implementation.model.RowData;
import dataimporter.implementation.model.TableData;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import dataimporter.implementation.utils.DataImporterUtils;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;


public class CsvDataReader {
    public TableData readCSVData(File csvFile, String delimiter, Character quoteCharacter, Character escapeCharacter, boolean addHeaderRow, List<String> headerNames) throws DataImporterRuntimeException {
        if (csvFile == null || !csvFile.exists()) {
            DataImporterUtils.logNode.error("CSV file not found.");
            throw new DataImporterRuntimeException("CSV file not found.");
        }

        var csvFileFormat = getCsvFormat(delimiter, quoteCharacter, escapeCharacter);
        List<RowData> rowDataList = new ArrayList<>();
        var headerRow = new RowData();
        var rowCounter = 1;
        try (var csvFileParser = CSVParser.parse(csvFile, StandardCharsets.UTF_8, csvFileFormat)) {
            for (CSVRecord csvRecord : csvFileParser) {
                List<CellData> cellDataList = createCellDataList(addHeaderRow, headerRow, csvRecord);
                if (csvRecord.getRecordNumber() == 1) {
                    checkHeaders(headerNames, cellDataList);
                    //Adds cell data to header list
                    headerRow.setCellData(cellDataList);
                    headerRow.setRowIndex(1); //RowIndex will be always 1
                    //If addHeaderRow is True, Add 1st record in row list as well
                    if (addHeaderRow) {
                        createRowList(rowDataList, csvRecord, cellDataList);
                    }
                } else {
                    createRowList(rowDataList, csvRecord, cellDataList);
                }
                rowCounter++;
            }
            var tableData = new TableData();
            tableData.setTableIndex(0);
            tableData.setRowDataList(rowDataList);

            return tableData;
        } catch (Exception ex) {
            DataImporterUtils.logNode.error("Error while parsing csv file at row no " + rowCounter, ex);
            throw new DataImporterRuntimeException("Error while parsing csv file at row no " + rowCounter, ex);
        }
    }

    private void checkHeaders(List<String> headerNames, List<CellData> cellDataList) {
        if (headerNames.size() > cellDataList.size()) {
            DataImporterUtils.logNode.error("Some of the columns are missing in uploaded file. Please upload file having same structure as selected Data Importer document.");
            throw new DataImporterRuntimeException("Some of the columns are missing in uploaded file. Please upload file having same structure as selected Data Importer document.");
        }
        Set<String> columnHeaderUploaded = cellDataList.stream()
                .map(CellData::getColumnHeader)
                .map(Object::toString)
                .collect(Collectors.toSet());

        for (String headerName : headerNames) {
            if (!columnHeaderUploaded.contains(headerName)) {
                DataImporterUtils.logNode.error("'" + headerName + "' column does not exist in sheet. Please upload file having same structure as selected Data Importer document.");
                throw new DataImporterRuntimeException("'" + headerName + "' column does not exist in sheet. Please upload file having same structure as selected Data Importer document.");
            }
        }
    }

    private List<CellData> createCellDataList(boolean addHeaderRow, RowData headerRow, CSVRecord csvRecord) {
        List<CellData> cellDataList = new ArrayList<>();
        for (var i = 0; i < csvRecord.size(); ++i) {
            CellData cellData;
            if (csvRecord.getRecordNumber() == 1) {
                //If addHeader is True, add column index as header value else actual column header value provided
                var headerName = csvRecord.get(i) ==  null ? csvRecord.get(i) : csvRecord.get(i).trim();
                cellData = new CellData(i, addHeaderRow ? String.valueOf(i + 1) : headerName, csvRecord.get(i), null);
            } else {
                cellData = new CellData(i, headerRow.getCellData().get(i).getColumnHeader(), csvRecord.get(i), null);
            }
            cellDataList.add(cellData);
        }
        return cellDataList;
    }

    private void createRowList(List<RowData> rowDataList, CSVRecord csvRecord, List<CellData> cellDataList) {
        var rowData = new RowData();
        rowData.setCellData(cellDataList);
        rowData.setRowIndex(csvRecord.getRecordNumber());
        rowDataList.add(rowData);
    }

    private CSVFormat getCsvFormat(String delimiter, Character quoteCharacter, Character escapeCharacter) {
        //If same quote and escape char used, Apache Commons CSV gives error to incorporate this we are not setting escape char in CSVFormat
        if(quoteCharacter.equals(escapeCharacter)){
            return CSVFormat.DEFAULT.builder().setDelimiter(delimiter).setQuote(quoteCharacter).setSkipHeaderRecord(false)
                    .setAllowMissingColumnNames(true) //if column name is blank => throws exception, to avoid exception, updated setAllowMissingColumnNames(true)
                    .setNullString("")
                    .build();

        }
        return CSVFormat.DEFAULT.builder().setDelimiter(delimiter).setQuote(quoteCharacter).setEscape(escapeCharacter).setSkipHeaderRecord(false)
                .setAllowMissingColumnNames(true) //if column name is blank => throws exception, to avoid exception, updated setAllowMissingColumnNames(true)
                .setNullString("")
                .build();
    }
}

