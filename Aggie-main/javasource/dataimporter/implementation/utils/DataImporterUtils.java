package dataimporter.implementation.utils;

import com.mendix.core.Core;
import com.mendix.logging.ILogNode;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IEntityProxy;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import com.mendix.systemwideinterfaces.core.meta.IMetaPrimitive;
import com.mendix.thirdparty.org.json.JSONObject;
import com.mendix.thirdparty.org.json.JSONStringer;
import dataimporter.implementation.enums.FileExtension;
import dataimporter.implementation.enums.FileType;
import dataimporter.implementation.model.CellData;
import dataimporter.implementation.model.ImportMappingParameters;
import dataimporter.proxies.ColumnAttributeMapping;
import dataimporter.proxies.CsvSheet;
import dataimporter.proxies.DataImporterElement;
import dataimporter.proxies.ExcelSheet;
import dataimporter.proxies.Template;
import dataimporter.proxies.constants.Constants;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.ExcelNumberFormat;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.RecordFormatException;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class DataImporterUtils {
    public static final ILogNode logNode = Core.getLogger(Constants.getLogNode());
    public static final String MISMATCH_DATA_TYPE_FOUND = "Mismatched data type found between cell and entity attribute.";
    // Format to match the expected date format in JSON
    public static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");

    private DataImporterUtils() {

    }

    public static FileExtension getFileExtension(String fileName) {
        final int lastdot = fileName.lastIndexOf(".");
        if (lastdot < 0) {
            logNode.error("Found file has no extension to derive format from.");
            throw new DataImporterRuntimeException("Found file has no extension to derive format from.");
        }
        switch (fileName.substring(lastdot)) {
            case ".xls":
                return FileExtension.XLS;
            case ".xlsx":
                return FileExtension.XLSX;
            case ".csv":
                return FileExtension.CSV;
            default:
                return FileExtension.UNKNOWN;
        }
    }

    public static JSONObject getTemplateJsonObject(FileType type, String MappingTemplate) {
        var templateMetaData = new JSONObject(MappingTemplate);
        var fileType = (templateMetaData.getInt("fileType"));
        if (type.ordinal() != fileType) {
            DataImporterUtils.logNode.error("Uploaded file type does not match with the file used to create selected Data Importer document.");
            throw new DataImporterRuntimeException("Uploaded file type does not match with the file used to create selected Data Importer document.");
        }
        return templateMetaData;
    }

    public static void deleteTempFile(File file) {
        if (file != null) {
            try {
                Files.delete(file.toPath());
            } catch (final Exception ignored) {
                DataImporterUtils.logNode.error("Could not delete temp data file.");
            }
        }
    }

    public static void handleSpecificExceptions(Exception e) {
        if (e instanceof OLE2NotOfficeXmlFileException) {
            DataImporterUtils.logNode.error("Document could not be imported. Please make sure the data file is valid and has the correct extension.");
            throw new DataImporterRuntimeException("Document could not be imported. Please make sure the data file is valid and has the correct extension.");
        } else if (e instanceof NotOfficeXmlFileException) {
            DataImporterUtils.logNode.error("Document could not be imported because this data file is not XLS or XLSX or CSV. Please make sure the data file is valid and has the correct extension.");
            throw new DataImporterRuntimeException("Document could not be imported because this data file is not XLS or XLSX or CSV. Please make sure the data file is valid and has the correct extension.");
        } else if (e instanceof RecordFormatException) {
            DataImporterUtils.logNode.error("Document could not be imported because one of its cell values is invalid or cannot be read.");
            throw new DataImporterRuntimeException("Document could not be imported because one of its cell values is invalid or cannot be read.");
        } else if (e instanceof EncryptedDocumentException) {
            DataImporterUtils.logNode.error("Document could not be imported because it is encrypted.");
            throw new DataImporterRuntimeException("Document could not be imported because it is encrypted.");
        }
    }

    public static String getColumnHeader(Cell cell, List<CellData> headerRowData, int columnIndex) {
        if (headerRowData == null) {
            return getCellName(cell);
        } else {
            var headerCellData = headerRowData.stream().filter(h -> h.getColumnIndex() == columnIndex).findFirst().orElse(null);
            return headerCellData == null ? null : headerCellData.getFormattedData().toString();
        }
    }

    public static boolean isCellDateFormatted(short dataFormat, String formatString) {
        var excelNoFormat = new ExcelNumberFormat(dataFormat, formatString);
        return DateUtil.isADateFormat(excelNoFormat);
    }

    public static String getCellName(Cell cell) {
        return CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1);
    }

    public static String sanitizeName(String name) {
        //Applying library conversion logic except reserved keywords
        name = name.replaceAll("[^a-zA-Z0-9_ ]+", "");
        name = name.replaceAll("[\\s\\xa0]+", " ").trim();
        return name.replaceAll("\\W+", "_");
    }

    public static Map<String, IMetaPrimitive> getIMetaPrimitiveMap(List<ColumnAttributeMapping> columnAttributeMappingMendixObjects) {
        Map<String, IMetaPrimitive> metaPrimitiveMap = new HashMap<>();
        for (ColumnAttributeMapping attributeMapping : columnAttributeMappingMendixObjects) {
            String attributeName = Core.getMetaPrimitive(attributeMapping.getAttribute()).getName();
            var iMetaPrimitive = Core.getMetaPrimitive(attributeMapping.getAttribute());
            metaPrimitiveMap.put(attributeName, iMetaPrimitive);
        }
        return metaPrimitiveMap;
    }

    public static IMendixObject processRowData(IContext context, List<CellData> dataRow, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects, FileType fileType) {
        // Store MetaPrimitives in a map to avoid multiple calls to Core API functions
        Map<String, IMetaPrimitive> metaPrimitiveMap = getIMetaPrimitiveMap(columnAttributeMappingMendixObjects);

        // Create a HashMap to store the CellData objects for each column header
        Map<String, CellData> cellDataMap = new HashMap<>();
        for (CellData cellData : dataRow) {
            cellDataMap.put(cellData.getColumnHeader(), cellData);
        }
        // Create the entity object
        IMendixObject entityObject = Core.instantiate(context, Core.getMetaPrimitive(columnAttributeMappingMendixObjects.get(0).getAttribute()).getParent().getName());
        for (ColumnAttributeMapping attributeMapping : columnAttributeMappingMendixObjects) {
            String attributeName = Core.getMetaPrimitive(attributeMapping.getAttribute()).getName();
            var cellData = cellDataMap.get(attributeMapping.getColumnName());
            if (cellData != null) {
                var iMetaPrimitive = metaPrimitiveMap.get(attributeName);
                try {
                    entityObject.setValue(context, attributeName, getMendixTypeObject(iMetaPrimitive.getType(), cellData, fileType));
                } catch (Exception e) {
                    throw new DataImporterRuntimeException(e.getMessage(), e);
                }
            }
        }
        return entityObject;
    }

    public static Map<String, Object> processRowDataJsonMapping(List<CellData> dataRow, List<DataImporterElement> dataImporterElementsMendixObjects, FileType fileType) {
        // Create a HashMap to store the CellData objects for each column header
        Map<String, CellData> cellDataMap = new HashMap<>();
        for (CellData cellData : dataRow) {
            cellDataMap.put(cellData.getColumnHeader(), cellData);
        }
        Map<String, Object> dataMap = new HashMap<>();
        for (DataImporterElement dataImporterElement : dataImporterElementsMendixObjects) {
            String columnHeader = dataImporterElement.getDecodedPath().trim();
            var cellData = cellDataMap.get(columnHeader);
            if (cellData != null) {
                var primitiveType = IMetaPrimitive.PrimitiveType.valueOf(dataImporterElement.getPrimitiveType());
                try {
                    dataMap.put(columnHeader, getMendixTypeObject(primitiveType, cellData, fileType));
                } catch (Exception e) {
                    throw new DataImporterRuntimeException(e.getMessage(), e);
                }
            }
        }
        return dataMap;
    }

    public static List<IMendixObject> callIntegrationImportStream(IContext context, List<Map<String, Object>> dataMapList, ImportMappingParameters mappingParameters) {
        String jsonString = getJsonString(dataMapList);
        try (InputStream inputStream = new ByteArrayInputStream(jsonString.getBytes(StandardCharsets.UTF_8))) {
            return Core.integration().importStream(
                    context,
                    inputStream,
                    mappingParameters.getImportMappingName(),
                    mappingParameters.getActionWhenNoObjectFound(),
                    null,
                    mappingParameters.getLimit(),
                    mappingParameters.getShouldCommit(),
                    false);
        } catch (IOException e) {
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
    }

    private static String getJsonString(List<Map<String, Object>> dataMapList) {
        // JSONStringer to build the entire JSON structure
        JSONStringer jsonStringer = new JSONStringer();
        jsonStringer.array(); // Begin JSON array
        for (Map<String, Object> dataMap : dataMapList) {
            jsonStringer.object(); // Start JSON object
            for (Map.Entry<String, Object> entry : dataMap.entrySet()) {
                jsonStringer.key(entry.getKey()) // Add key
                        .value(entry.getValue() instanceof Date ? simpleDateFormat.format(entry.getValue()) : entry.getValue()); // Add value
            }
            jsonStringer.endObject(); // End JSON object
        }
        jsonStringer.endArray(); // End JSON array
        return jsonStringer.toString();
    }

    public static Object getMendixTypeObject(IMetaPrimitive.PrimitiveType primitiveType, CellData cellData, FileType fileType) {
        if (fileType.equals(FileType.EXCEL) && cellData.getFormattedData() == null) {
            return null;
        }
        if (fileType.equals(FileType.CSV) && cellData.getRawData() == null) {
            return null;
        }

        try {
            switch (primitiveType) {
                case String:
                case Boolean:
                case DateTime:
                case Decimal:
                case Integer:
                case Long:
                    return DataTypeConversionUtils.getConvertedData(cellData, fileType, primitiveType);
                default: {
                    logNode.error(MISMATCH_DATA_TYPE_FOUND);
                    return new DataReaderException(MISMATCH_DATA_TYPE_FOUND);
                }
            }
        } catch (Exception e) {
            throw new DataImporterRuntimeException(e.getMessage(), e);
        }
    }

    public static List<IMendixObject> retrieveByPath(IContext context, IMendixObject object, String memberName) {
        return Core.retrieveByPath(context, object, memberName);
    }

    public static void processChildElements(IContext context, IMendixObject parentElement, List<DataImporterElement> dataImporterElements) {
        // Retrieve child elements of the current data importer element
        List<IMendixObject> dataImporterElementChildObjects = Core.retrieveByPath(context, parentElement, DataImporterElement.MemberNames.Parent.toString(), true);
        for (IMendixObject childElement : dataImporterElementChildObjects) {
            // Retrieve child elements of the child and initialize them
            List<IMendixObject> jsonObjectChildObjects = Core.retrieveByPath(context, childElement, DataImporterElement.MemberNames.Parent.toString(), true);
            for (IMendixObject jsonObjectChildElement : jsonObjectChildObjects) {
                dataImporterElements.add(DataImporterElement.initialize(context, jsonObjectChildElement));
            }
        }
    }

    public static Map<IEntityProxy, List<DataImporterElement>> getSheetElementMap(IContext context, IMendixObject mappingTemplate) {
        Map<IEntityProxy, List<DataImporterElement>> sheetDataImporterElementMap = new HashMap<>();
        // Retrieve all template sheets
        List<IMendixObject> templateSheets = Core.retrieveByPath(context, mappingTemplate, ExcelSheet.MemberNames.ExcelSheet_Template.toString());
        // Process each template sheet
        for (IMendixObject templateSheetObject : templateSheets) {
            List<DataImporterElement> dataImporterElements = new ArrayList<>();
            // Retrieve and process all data importer elements
            List<IMendixObject> dataImporterElementObjects = DataImporterUtils.retrieveByPath(context, templateSheetObject, DataImporterElement.MemberNames.DataImporterElement_ExcelSheet.toString());
            for (IMendixObject dataImporterElementMendixObj : dataImporterElementObjects) {
                // Retrieve child elements and add to the list
                DataImporterUtils.processChildElements(context, dataImporterElementMendixObj, dataImporterElements);
            }
            // Add result to the map
            sheetDataImporterElementMap.put(ExcelSheet.initialize(context, templateSheetObject), dataImporterElements);
        }
        return sheetDataImporterElementMap;
    }

    public static void validateJsonMappingHeaders(List<DataImporterElement> dataImporterElementsMendixObjects, List<CellData> headerRowData, String sheetName) {
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

        for (DataImporterElement dataImporterElement : dataImporterElementsMendixObjects) {
            String columnName = dataImporterElement.getDecodedPath().trim();
            if (!headerColumnNames.contains(columnName)) {
                var errorMsg = String.format("Column with name: '%s' is not found in sheet: '%s'", columnName, sheetName);
                DataImporterUtils.logNode.error(errorMsg);
                throw new DataImporterRuntimeException(errorMsg);
            }
        }
    }

    public static File getFile(IContext context, IMendixObject excelDocument) {
        final var file = new File(Core.getConfiguration().getTempPath().getAbsolutePath() + "/Mendix_DataImporter_" + excelDocument.getId().toLong(), "");
        try (var inputStream = Core.getFileDocumentContent(context, excelDocument);
             OutputStream outputstream = new FileOutputStream(file)) {
            final var buffer = new byte[4 * 1024];
            int length;
            while ((length = inputStream.read(buffer)) > 0) {
                outputstream.write(buffer, 0, length);
            }
        } catch (IOException e) {
            DataImporterUtils.logNode.error("You must upload a file document before the the data can be imported.");
            throw new DataImporterRuntimeException("You must upload a file document before the the data can be imported.");
        }
        return file;
    }

    public static Template getTemplateMendixObjectFromJSON(IContext context, String MappingTemplate, FileType fileType) {
        var templateMetaData = DataImporterUtils.getTemplateJsonObject(fileType, MappingTemplate);

        var template = new Template(context);
        template.setTemplateName(templateMetaData.getString("templateName"));
        var entityMetaDataArr = templateMetaData.getJSONArray("entityMetaData");
        Object sheetObj = null;

        var entityNameList = new ArrayList<>();
        for (var i = 0; i < entityMetaDataArr.length(); i++) {
            var sheetMetaData = entityMetaDataArr.getJSONObject(i);
            if(!sheetMetaData.isNull("excelMapping")){
                var sheet = new ExcelSheet(context);
                var excelMappingData = sheetMetaData.getJSONObject("excelMapping");
                sheet.setSheetName(excelMappingData.getString("excelSheetName"));
                sheet.setHeaderRowStartsAt(excelMappingData.getInt("headerRowNo"));
                sheet.setDataRowStartsAt(excelMappingData.getInt("readDataFrom"));
                sheet.setExcelSheet_Template(template);
                sheetObj = sheet;
            }
            if(!sheetMetaData.isNull("csvMapping")){
                var csvMappingData = sheetMetaData.getJSONObject("csvMapping");
                var sheet = new CsvSheet(context);
                sheet.setDelimiter(csvMappingData.getString("delimiter"));
                sheet.setQuoteCharacter(csvMappingData.getString("quoteCharacter"));
                sheet.setAddHeaderRow(csvMappingData.getBoolean("addHeaderRow"));
                sheet.setEscapeCharacter(csvMappingData.getString("escapeCharacter"));
                sheet.setCsvSheet_Template(template);
                sheetObj = sheet;
            }

            entityNameList.add(sheetMetaData.getString("entityName"));

            var columnAttributeMapping = sheetMetaData.getJSONArray("columnMetadata");

            for (var j = 0; j < columnAttributeMapping.length(); j++) {
                var columnData = columnAttributeMapping.getJSONObject(j);
                var attributeMapping = new ColumnAttributeMapping(context);
                setAttributeAndColumnName(columnData, attributeMapping);
                if(sheetObj instanceof ExcelSheet){
                    attributeMapping.setColumnAttributeMapping_ExcelSheet((ExcelSheet) sheetObj);
                }
                if(sheetObj instanceof CsvSheet){
                    attributeMapping.setColumnAttributeMapping_CsvSheet((CsvSheet) sheetObj);
                }
            }

            if(!sheetMetaData.isNull("dataImporterElement")){
                processDataImporterElement(context,sheetMetaData, sheetObj);
            }
        }
        /*if (!entityNameList.contains(this.Entity)) {
            throw new CoreException("The entity '" + this.Entity + "' is not present in the '" + template.getTemplateName() + "'.");
        }*/
        return template;
    }
    private static void setAttributeAndColumnName(JSONObject columnData, ColumnAttributeMapping attributeMapping) {
        // Remove below conditions and keep only else code when we upgrade to platform version 10.7.0+ DHC-1754
        if (columnData.has("name")) {
            attributeMapping.setAttribute(columnData.getString("name"));
        } else {
            attributeMapping.setAttribute(columnData.getString("attributeName"));
        }
        if (columnData.has("columnName")) {
            attributeMapping.setColumnName(columnData.getString("columnName"));
        } else {
            attributeMapping.setColumnName(columnData.getString("originalName"));
        }
    }

    private static void processDataImporterElement(IContext context,JSONObject sheetMetaDataJsonObj, Object sheet) {
        var rootDataImporterElement = sheetMetaDataJsonObj.getJSONObject("dataImporterElement");
        if(rootDataImporterElement == null){
            return;
        }
        var rootDataImporterElementProxyObj =  setDataImporterElementAttributesFromJSON(context,rootDataImporterElement);
        var childrenDataImporterElements = rootDataImporterElement.getJSONArray("childrens");
        for (var j = 0; j < childrenDataImporterElements.length(); j++) {
            var childrenJsonObj = childrenDataImporterElements.getJSONObject(j);
            var childrenElementProxyObj = setDataImporterElementAttributesFromJSON(context,childrenJsonObj);
            childrenElementProxyObj.setParent(rootDataImporterElementProxyObj);

            var columnRootElements = childrenJsonObj.getJSONArray("childrens");
            for (var k = 0; k < columnRootElements.length(); k++) {
                var columnElementJsonObj = columnRootElements.getJSONObject(k);
                var columnElementProxyObj = setDataImporterElementAttributesFromJSON(context,columnElementJsonObj);
                columnElementProxyObj.setParent(childrenElementProxyObj);
            }
        }
        if (sheet instanceof ExcelSheet) {
            rootDataImporterElementProxyObj.setDataImporterElement_ExcelSheet((ExcelSheet) sheet);
        } else if (sheet instanceof CsvSheet) {
            rootDataImporterElementProxyObj.setDataImporterElement_CsvSheet((CsvSheet) sheet);
        }
    }

    private static DataImporterElement setDataImporterElementAttributesFromJSON(IContext context,JSONObject elementJsonObj) {
        var dataImporterElementProxyObj = new DataImporterElement(context);
        dataImporterElementProxyObj.setElementType(elementJsonObj.getInt("elementType"));
        dataImporterElementProxyObj.setPrimitiveType(elementJsonObj.getString("primitiveType"));
        dataImporterElementProxyObj.setPath(elementJsonObj.getString("path"));
        dataImporterElementProxyObj.setDecodedPath(elementJsonObj.getString("decodedPath"));
        dataImporterElementProxyObj.setIsDefaultType(elementJsonObj.getBoolean("isDefaultType"));
        dataImporterElementProxyObj.setMinOccurs(elementJsonObj.getInt("minOccurs"));
        dataImporterElementProxyObj.setMaxOccurs(elementJsonObj.getInt("maxOccurs"));
        dataImporterElementProxyObj.setNillable(elementJsonObj.getBoolean("nillable"));
        dataImporterElementProxyObj.setExposedName(elementJsonObj.getString("exposedName"));
        dataImporterElementProxyObj.setExposedItemName(elementJsonObj.getString("exposedItemName"));
        dataImporterElementProxyObj.setMaxLength(elementJsonObj.getInt("maxLength"));
        dataImporterElementProxyObj.setFractionDigits(elementJsonObj.getInt("fractionDigits"));
        dataImporterElementProxyObj.setTotalDigits(elementJsonObj.getInt("totalDigits"));
        dataImporterElementProxyObj.setErrorMessage(elementJsonObj.getString("errorMessage"));
        dataImporterElementProxyObj.setWarningMessage(elementJsonObj.getString("warningMessage"));
        dataImporterElementProxyObj.setOriginalValue(elementJsonObj.getString("originalValue"));
        return dataImporterElementProxyObj;
    }
}
