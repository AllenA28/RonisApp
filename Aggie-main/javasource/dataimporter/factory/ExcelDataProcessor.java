package dataimporter.factory;

import com.mendix.core.Core;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IEntityProxy;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import dataimporter.implementation.enums.FileType;
import dataimporter.implementation.model.CellData;
import dataimporter.implementation.model.ImportMappingParameters;
import dataimporter.implementation.service.ExcelDataReader;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import dataimporter.implementation.utils.DataImporterUtils;
import dataimporter.implementation.utils.DataReaderException;
import dataimporter.proxies.ColumnAttributeMapping;
import dataimporter.proxies.DataImporterElement;
import dataimporter.proxies.ExcelSheet;

import java.io.File;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelDataProcessor implements IDataProcessor {
    static final String STARTED = " started.";
    static final String FROM_SHEET = " from sheet ";
    private String sheetName;

    @Override
    public Map<IEntityProxy, List<ColumnAttributeMapping>> startImport(IContext context, IMendixObject mappingTemplate) {
        Map<IEntityProxy, List<ColumnAttributeMapping>> sheetColumnMappingMap = new HashMap<>();
        List<IMendixObject> templateSheets = Core.retrieveByPath(context, mappingTemplate, ExcelSheet.MemberNames.ExcelSheet_Template.toString());
        for (IMendixObject templateSheetObject : templateSheets) {
            List<ColumnAttributeMapping> columnAttributeMappings = new ArrayList<>();
            List<IMendixObject> columnAttributeMappingObjects = Core.retrieveByPath(context, templateSheetObject, ColumnAttributeMapping.MemberNames.ColumnAttributeMapping_ExcelSheet.toString());
            for (IMendixObject columnAttributeMapping : columnAttributeMappingObjects) {
                columnAttributeMappings.add(ColumnAttributeMapping.initialize(context, columnAttributeMapping));
            }
            sheetColumnMappingMap.put(ExcelSheet.initialize(context, templateSheetObject), columnAttributeMappings);
        }
        return sheetColumnMappingMap;
    }

    @Override
    public List<IMendixObject> parseData(IContext context, File file, IEntityProxy sheetMendixObject, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects) {
        List<IMendixObject> importerList = new ArrayList<>();
        ExcelSheet sheetExcel = (ExcelSheet) sheetMendixObject;
        int dataRowIndex = sheetExcel.getDataRowStartsAt() - 1;
        try (var excelDataReader = new ExcelDataReader(file)) {
            sheetName = sheetExcel.getSheetName();
            excelDataReader.openSheet(sheetName);
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel header row from sheet: '" + sheetName + "'" + STARTED);
            }
            List<CellData> headerRowData = excelDataReader.readHeaderRow(sheetExcel.getHeaderRowStartsAt() - 1);
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel header row from sheet: '" + sheetName + "' finished. Found '" + headerRowData.size() + "' columns.");
            }
            if (headerRowData == null || headerRowData.isEmpty()) {
                DataImporterUtils.logNode.error("No column information could be found in sheet: '" + sheetName + "'");
                throw new DataImporterRuntimeException("No column information could be found in sheet: '" + sheetName + "'");
            }
            Set<String> headerColumnNames = headerRowData.stream()
                    .map(cell -> cell.getFormattedData().toString())
                    .collect(Collectors.toSet());
            for (ColumnAttributeMapping columnAttributeMapping : columnAttributeMappingMendixObjects) {
                if (!headerColumnNames.contains(columnAttributeMapping.getColumnName())) {
                    var errorMsg = String.format("Column with name: '%s' is not found in sheet: '%s'", columnAttributeMapping.getColumnName(), sheetName);
                    DataImporterUtils.logNode.error(errorMsg);
                    throw new DataImporterRuntimeException(errorMsg);
                }
            }
            int totalRowCount = excelDataReader.getNumberOfRows();
            for (int currentRowIndex = dataRowIndex; currentRowIndex <= totalRowCount; currentRowIndex++) {
                List<CellData> excelCellDataList = readExcelRow(excelDataReader, headerRowData, currentRowIndex);
                if (!excelCellDataList.isEmpty()) {
                    importerList.add(DataImporterUtils.processRowData(context, excelCellDataList, columnAttributeMappingMendixObjects, FileType.EXCEL));
                }
            }
        } catch (Exception e) {
            DataImporterUtils.logNode.error("An error occurred while processing data.");
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
        return importerList;
    }

    private List<CellData> readExcelRow(ExcelDataReader excelDataReader, List<CellData> headerRowData, int currentRowIndex) throws DataReaderException {
        try {
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel row: " + currentRowIndex + 1 + FROM_SHEET + sheetName + STARTED);
            }
            List<CellData> cellDataList = excelDataReader.readDataRow(currentRowIndex, headerRowData);
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel row: " + currentRowIndex + 1 + FROM_SHEET + sheetName + " finished. Found " + cellDataList.size() + " cells.");
            }
            return cellDataList;
        } catch (Exception e) {
            var errorMsg = String.format("Unable to import sheet row '%d' %s '%s'", currentRowIndex + 1, FROM_SHEET, sheetName);
            DataImporterUtils.logNode.error(errorMsg, e);
            throw new DataReaderException(errorMsg, e);
        }
    }

    public Map<IEntityProxy, List<DataImporterElement>> startJsonMappingImport(IContext context, IMendixObject mappingTemplate) {
        return DataImporterUtils.getSheetElementMap(context, mappingTemplate);
    }


    @Override
    public List<IMendixObject> parseJsonMappingData(IContext context, File excelFile, IEntityProxy sheetMendixObject, List<DataImporterElement> dataImporterElementsMendixObjects, ImportMappingParameters mappingParameters) {
        List<Map<String, Object>> jsonStream = new ArrayList<>();
        ExcelSheet sheetExcel = (ExcelSheet) sheetMendixObject;
        List<IMendixObject> integrationImportOutput;
        int dataRowIndex = sheetExcel.getDataRowStartsAt() - 1;
        try (var excelDataReader = new ExcelDataReader(excelFile)) {
            sheetName = mappingParameters.getSheetName();
            excelDataReader.openSheet(sheetName);
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel header row from sheet: '" + sheetName + "'" + STARTED);
            }
            List<CellData> headerRowData = excelDataReader.readHeaderRow(sheetExcel.getHeaderRowStartsAt() - 1);
            DataImporterUtils.validateJsonMappingHeaders(dataImporterElementsMendixObjects, headerRowData, sheetName);
            int totalRowCount = excelDataReader.getNumberOfRows();
            for (int currentRowIndex = dataRowIndex; currentRowIndex <= totalRowCount; currentRowIndex++) {
                List<CellData> excelCellDataList = readExcelRow(excelDataReader, headerRowData, currentRowIndex);
                if (!excelCellDataList.isEmpty()) {
                    jsonStream.add(DataImporterUtils.processRowDataJsonMapping(excelCellDataList, dataImporterElementsMendixObjects, FileType.EXCEL));
                }
            }
            integrationImportOutput = DataImporterUtils.callIntegrationImportStream(context, jsonStream, mappingParameters);
        } catch (Exception e) {
            DataImporterUtils.logNode.error("An error occurred while processing data.");
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
        return integrationImportOutput;
    }
}
