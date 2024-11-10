package dataimporter.factory;

import com.mendix.core.Core;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IEntityProxy;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import dataimporter.implementation.enums.FileType;
import dataimporter.implementation.model.CellData;
import dataimporter.implementation.model.ImportMappingParameters;
import dataimporter.implementation.service.ExcelXLSXDataReader;
import dataimporter.implementation.utils.DataImporterRuntimeException;
import dataimporter.implementation.utils.DataImporterUtils;
import dataimporter.proxies.ColumnAttributeMapping;
import dataimporter.proxies.DataImporterElement;
import dataimporter.proxies.ExcelSheet;

import java.io.File;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelXlsxDataProcessor implements IDataProcessor {
    static final String STARTED = " started.";
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
    public List<IMendixObject> parseData(IContext context, File excelFile, IEntityProxy sheetMendixObject, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects) {
        List<IMendixObject> importedDataList = new ArrayList<>();
        ExcelSheet sheetExcel = (ExcelSheet) sheetMendixObject;
        int startRowIndex = sheetExcel.getDataRowStartsAt() - 1;
        String sheetName = sheetExcel.getSheetName();
        try {
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel header row from sheet: '" + sheetName + "'" + STARTED);
            }
            List<CellData> headerRowData = ExcelXLSXDataReader.readHeaderRow(excelFile, sheetName, sheetExcel.getHeaderRowStartsAt() - 1);
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel header row from sheet: '" + sheetName + "' finished. Found '" + headerRowData.size() + "' columns.");
            }
            if (headerRowData.isEmpty()) {
                var errorMsg = String.format("No column information could be found in sheet: '%s'", sheetName);
                DataImporterUtils.logNode.error(errorMsg);
                throw new DataImporterRuntimeException(errorMsg);
            }
            Set<String> headerColumnNames = headerRowData.stream()
                    .map(cell -> cell.getFormattedData().toString())
                    .filter(Objects::nonNull)
                    .collect(Collectors.toSet());

            for (ColumnAttributeMapping columnAttributeMapping : columnAttributeMappingMendixObjects) {
                if (!headerColumnNames.contains(columnAttributeMapping.getColumnName())) {
                    var errorMsg = String.format("Column with name: '%s' is not found in sheet: '%s'", columnAttributeMapping.getColumnName(), sheetName);
                    DataImporterUtils.logNode.error(errorMsg);
                    throw new DataImporterRuntimeException(errorMsg);
                }
            }
            List<List<CellData>> excelCellDataList = ExcelXLSXDataReader.readDataRow(excelFile, sheetName, startRowIndex, headerRowData);
            for (List<CellData> rowData : excelCellDataList) {
                if (!rowData.isEmpty()) {
                    importedDataList.add(DataImporterUtils.processRowData(context, rowData, columnAttributeMappingMendixObjects, FileType.EXCEL));
                }
            }
        } catch (Exception e) {
            DataImporterUtils.logNode.error("An error occurred while processing data.");
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
        return importedDataList;
    }

    public Map<IEntityProxy, List<DataImporterElement>> startJsonMappingImport(IContext context, IMendixObject mappingTemplate) {
        return DataImporterUtils.getSheetElementMap(context, mappingTemplate);
    }

    @Override
    public List<IMendixObject> parseJsonMappingData(IContext context, File excelFile, IEntityProxy sheetMendixObject, List<DataImporterElement> dataImporterElementsMendixObjects, ImportMappingParameters mappingParameters) {
        List<Map<String, Object>> jsonStream = new ArrayList<>();
        ExcelSheet sheetExcel = (ExcelSheet) sheetMendixObject;
        List<IMendixObject> integrationImportOutput;
        int startRowIndex = sheetExcel.getDataRowStartsAt() - 1;
        String sheetName = mappingParameters.getSheetName();
        try {
            if (DataImporterUtils.logNode.isTraceEnabled()) {
                DataImporterUtils.logNode.trace("Reading excel header row from sheet: '" + sheetName + "'" + STARTED);
            }
            List<CellData> headerRowData = ExcelXLSXDataReader.readHeaderRow(excelFile, sheetName, sheetExcel.getHeaderRowStartsAt() - 1);

            DataImporterUtils.validateJsonMappingHeaders(dataImporterElementsMendixObjects, headerRowData, sheetName);

            List<List<CellData>> excelCellDataList = ExcelXLSXDataReader.readDataRow(excelFile, sheetName, startRowIndex, headerRowData);
            for (List<CellData> rowData : excelCellDataList) {
                if (!rowData.isEmpty()) {
                    jsonStream.add(DataImporterUtils.processRowDataJsonMapping(rowData, dataImporterElementsMendixObjects, FileType.EXCEL));
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
