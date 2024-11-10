package dataimporter.factory;

import com.mendix.core.Core;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IEntityProxy;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import dataimporter.implementation.enums.FileType;
import dataimporter.implementation.model.CellData;
import dataimporter.implementation.model.ImportMappingParameters;
import dataimporter.implementation.model.RowData;
import dataimporter.implementation.service.CsvDataReader;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import dataimporter.implementation.utils.DataImporterUtils;
import dataimporter.proxies.ColumnAttributeMapping;
import dataimporter.proxies.CsvSheet;
import dataimporter.proxies.DataImporterElement;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CsvDataProcessor implements IDataProcessor {

    public Map<IEntityProxy, List<ColumnAttributeMapping>> startImport(IContext context, IMendixObject mappingTemplate) {
        Map<IEntityProxy, List<ColumnAttributeMapping>> sheetColumnMappingMap = new HashMap<>();
        List<IMendixObject> templateSheets = Core.retrieveByPath(context, mappingTemplate, CsvSheet.MemberNames.CsvSheet_Template.toString());
        for (IMendixObject templateSheetObject : templateSheets) {
            List<ColumnAttributeMapping> columnAttributeMappings = new ArrayList<>();
            List<IMendixObject> columnAttributeMappingObjects = Core.retrieveByPath(context, templateSheetObject, ColumnAttributeMapping.MemberNames.ColumnAttributeMapping_CsvSheet.toString());
            for (IMendixObject columnAttributeMapping : columnAttributeMappingObjects) {
                columnAttributeMappings.add(ColumnAttributeMapping.initialize(context, columnAttributeMapping));
            }
            sheetColumnMappingMap.put(CsvSheet.initialize(context, templateSheetObject), columnAttributeMappings);
        }
        return sheetColumnMappingMap;
    }

    public List<IMendixObject> parseData(IContext context, File file, IEntityProxy sheetMendixObject, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects) {
        List<IMendixObject> importerList = new ArrayList<>();
        CsvSheet sheetCsv = (CsvSheet) sheetMendixObject;
        try {
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading csv file with configuration params, delimiter: '" + sheetCsv.getDelimiter() + "' escape character: '" + sheetCsv.getEscapeCharacter() + "' quote character: '" + sheetCsv.getQuoteCharacter() + "' addHeaderRow: '" + sheetCsv.getAddHeaderRow() + "'");
            }
            List<String> headerNames = new ArrayList<>();

            for (ColumnAttributeMapping columnAttributeMapping : columnAttributeMappingMendixObjects) {
                headerNames.add(columnAttributeMapping.getColumnName());
            }
            if (!sheetCsv.getEscapeCharacter().equals("") || !sheetCsv.getQuoteCharacter().equals("")) {
                var table = new CsvDataReader().readCSVData(file, sheetCsv.getDelimiter(), sheetCsv.getQuoteCharacter().charAt(0), sheetCsv.getEscapeCharacter().charAt(0), sheetCsv.getAddHeaderRow(), headerNames);
                var rowList = table.getRowDataList();

                for (RowData rowData : rowList) {
                    processRow(context, columnAttributeMappingMendixObjects, importerList, rowData);
                }
            } else {
                var errorMsg = "Unable to parse csv file due to invalid csv config parameters";
                DataImporterUtils.logNode.error(errorMsg);
                throw new DataImporterRuntimeException(errorMsg);
            }
        } catch (Exception e) {
            DataImporterUtils.logNode.error("An error occurred while parsing csv sheet.", e);
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
        return importerList;
    }

    private void processRow(IContext context, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects, List<IMendixObject> importerList, RowData rowData) {
        try {
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Processing data for row number : '" + rowData.getRowIndex() + "' in csv sheet");
            }
            List<CellData> csvCellDataList = rowData.getCellData();
            importerList.add(DataImporterUtils.processRowData(context, csvCellDataList, columnAttributeMappingMendixObjects, FileType.CSV));
        } catch (Exception ex) {
            DataImporterUtils.logNode.error("An error occurred while processing data for row number  " + rowData.getRowIndex() + " in csv sheet", ex);
            throw new DataImporterRuntimeException(ex.getMessage(), ex);
        }
    }

    @Override
    public Map<IEntityProxy, List<DataImporterElement>> startJsonMappingImport(IContext context, IMendixObject mappingTemplate) {
        Map<IEntityProxy, List<DataImporterElement>> sheetDataImporterElementMap = new HashMap<>();
        // Retrieve all template sheets
        List<IMendixObject> templateSheets = Core.retrieveByPath(context, mappingTemplate, CsvSheet.MemberNames.CsvSheet_Template.toString());
        // Process each template sheet
        for (IMendixObject templateSheetObject : templateSheets) {
            List<DataImporterElement> dataImporterElements = new ArrayList<>();
            // Retrieve and process all data importer elements
            List<IMendixObject> dataImporterElementObjects = DataImporterUtils.retrieveByPath(context, templateSheetObject, DataImporterElement.MemberNames.DataImporterElement_CsvSheet.toString());
            for (IMendixObject dataImporterElement : dataImporterElementObjects) {
                // Retrieve child elements and add to the list
                DataImporterUtils.processChildElements(context, dataImporterElement, dataImporterElements);
            }
            // Add result to the map
            sheetDataImporterElementMap.put(CsvSheet.initialize(context, templateSheetObject), dataImporterElements);
        }
        return sheetDataImporterElementMap;
    }

    @Override
    public List<IMendixObject> parseJsonMappingData(IContext context, File file, IEntityProxy sheetMendixObject, List<DataImporterElement> dataImporterElementsMendixObjects, ImportMappingParameters mappingParameters) {
        List<Map<String, Object>> jsonStream = new ArrayList<>();
        CsvSheet sheetCsv = (CsvSheet) sheetMendixObject;
        List<IMendixObject> integrationImportOutput;
        try {
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading csv file with configuration params, delimiter: '" + sheetCsv.getDelimiter() + "' escape character: '" + sheetCsv.getEscapeCharacter() + "' quote character: '" + sheetCsv.getQuoteCharacter() + "' addHeaderRow: '" + sheetCsv.getAddHeaderRow() + "'");
            }
            List<String> headerNames = new ArrayList<>();
            for (DataImporterElement dataImporterElement : dataImporterElementsMendixObjects) {
                headerNames.add(dataImporterElement.getDecodedPath().trim());
            }
            if (!sheetCsv.getEscapeCharacter().equals("") || !sheetCsv.getQuoteCharacter().equals("")) {
                var table = new CsvDataReader().readCSVData(file, sheetCsv.getDelimiter(), sheetCsv.getQuoteCharacter().charAt(0), sheetCsv.getEscapeCharacter().charAt(0), sheetCsv.getAddHeaderRow(), headerNames);
                var rowList = table.getRowDataList();
                for (RowData rowData : rowList) {
                    processRowJsonMapping(dataImporterElementsMendixObjects, jsonStream, rowData);
                }
                integrationImportOutput = DataImporterUtils.callIntegrationImportStream(context, jsonStream, mappingParameters);
            } else {
                var errorMsg = "Unable to parse csv file due to invalid csv config parameters";
                DataImporterUtils.logNode.error(errorMsg);
                throw new DataImporterRuntimeException(errorMsg);
            }
        } catch (Exception e) {
            DataImporterUtils.logNode.error("An error occurred while parsing csv sheet.", e);
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
        return integrationImportOutput;
    }

    private void processRowJsonMapping(List<DataImporterElement> dataImporterElementsMendixObjects, List<Map<String, Object>> jsonStream, RowData rowData) {
        try {
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Processing data for row number : '" + rowData.getRowIndex() + "' in csv sheet");
            }
            List<CellData> csvCellDataList = rowData.getCellData();
            jsonStream.add(DataImporterUtils.processRowDataJsonMapping(csvCellDataList, dataImporterElementsMendixObjects, FileType.CSV));
        } catch (Exception ex) {
            DataImporterUtils.logNode.error("An error occurred while processing data for row number  " + rowData.getRowIndex() + " in csv sheet", ex);
            throw new DataImporterRuntimeException(ex.getMessage(), ex);
        }
    }
}
