// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
// Special characters, e.g., é, ö, à, etc. are supported in comments.

package dataimporter.actions;

import com.mendix.core.CoreException;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IEntityProxy;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import com.mendix.webui.CustomJavaAction;
import dataimporter.factory.DataImporterFactory;
import dataimporter.implementation.enums.FileExtension;
import dataimporter.implementation.enums.FileType;
import dataimporter.implementation.utils.DataImporterUtils;
import dataimporter.proxies.ColumnAttributeMapping;
import dataimporter.proxies.Template;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.util.RecordFormatException;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class DataImport extends CustomJavaAction<java.util.List<IMendixObject>>
{
	/** @deprecated use DataFile.getMendixObject() instead. */
	@java.lang.Deprecated(forRemoval = true)
	private final IMendixObject __DataFile;
	private final system.proxies.FileDocument DataFile;
	private final java.lang.String MappingTemplate;
	private final java.lang.String Entity;

	public DataImport(
		IContext context,
		IMendixObject _dataFile,
		java.lang.String _mappingTemplate,
		java.lang.String _entity
	)
	{
		super(context);
		this.__DataFile = _dataFile;
		this.DataFile = _dataFile == null ? null : system.proxies.FileDocument.initialize(getContext(), _dataFile);
		this.MappingTemplate = _mappingTemplate;
		this.Entity = _entity;
	}

	@java.lang.Override
	public java.util.List<IMendixObject> executeAction() throws Exception
	{
		// BEGIN USER CODE
        if (this.DataFile == null || !this.DataFile.getHasContents() || this.DataFile.getSize() == 0) {
            DataImporterUtils.logNode.error("You must upload a valid file document before the data can be imported.");
            throw new CoreException("You must upload a valid file document before the data can be imported.");
        }
        if (this.MappingTemplate == null || this.MappingTemplate.isBlank()) {
            DataImporterUtils.logNode.error("Selected Data Importer document is empty.");
            throw new CoreException("Selected Data Importer document is empty.");
        }
        String fileName = this.DataFile.getName().toLowerCase(Locale.ROOT);
        if (DataImporterUtils.getFileExtension(fileName).equals(FileExtension.UNKNOWN)) {
            DataImporterUtils.logNode.error("Uploaded file type is not supported. Please upload supported file type.");
            throw new CoreException("Uploaded file type is not supported. Please upload supported file type.");
        }

        var fileType = FileType.EXCEL;

        if (DataImporterUtils.getFileExtension(fileName).equals(FileExtension.XLS) || DataImporterUtils.getFileExtension(fileName).equals(FileExtension.XLSX)) {
          fileType = FileType.EXCEL;
        }
        if (DataImporterUtils.getFileExtension(fileName).equals(FileExtension.CSV)) {
          fileType = FileType.CSV;
        }
        Template mappingTemplate = DataImporterUtils.getTemplateMendixObjectFromJSON(getContext(),this.MappingTemplate, fileType);//getExcelMendixObjectFromJSON();
       List<IMendixObject> importedList = new ArrayList<>();
        if (mappingTemplate != null) {
            Map<IEntityProxy, List<ColumnAttributeMapping>> sheetColumnMappingMap = DataImporterFactory.getDataProcessor(DataImporterUtils.getFileExtension(fileName)).startImport(this.getContext(), mappingTemplate.getMendixObject());
            importedList = getMendixObjectList(DataImporterUtils.getFile(this.getContext(), this.DataFile.getMendixObject()), fileName, sheetColumnMappingMap);
        }
        if (importedList == null) {
            DataImporterUtils.logNode.error("There is some problem occurred while processing the file");
            throw new CoreException("There is some problem occurred while processing the file");
        }
        return importedList;
		// END USER CODE
	}

	/**
	 * Returns a string representation of this action
	 * @return a string representation of this action
	 */
	@java.lang.Override
	public java.lang.String toString()
	{
		return "DataImport";
	}

	// BEGIN EXTRA CODE
    private List<IMendixObject> getMendixObjectList(java.io.File file, String fileName, Map<IEntityProxy, List<ColumnAttributeMapping>> sheetColumnMappingMap) throws CoreException {
        List<IMendixObject> importedList = null;
        for (Map.Entry<IEntityProxy, List<ColumnAttributeMapping>> entry : sheetColumnMappingMap.entrySet()) {
            importedList = parseSheetData(file, fileName, entry);
        }
        return importedList;
    }

    private List<IMendixObject> parseSheetData(File file, String fileName, Map.Entry<IEntityProxy, List<ColumnAttributeMapping>> entry) throws CoreException {
        final long importStartTime = System.nanoTime();
        final var ERROR_WHILE_IMPORTING = "Error while importing: '";
        final var MS_BECAUSE = " ms, because: ";
        List<IMendixObject> importedList = null;
        try {
            importedList = DataImporterFactory.getDataProcessor(DataImporterUtils.getFileExtension(fileName)).parseData(this.getContext(), file, entry.getKey(), entry.getValue());

        } catch (NotOfficeXmlFileException | RecordFormatException | EncryptedDocumentException e) {
            DataImporterUtils.logNode.error(ERROR_WHILE_IMPORTING + fileName + "' " + ((System.nanoTime() - importStartTime) / 1000000) + MS_BECAUSE + e.getMessage());
            DataImporterUtils.handleSpecificExceptions(e);
        } catch (Exception e) {
            DataImporterUtils.logNode.error(ERROR_WHILE_IMPORTING + fileName + "' " + ((System.nanoTime() - importStartTime) / 1000000) + MS_BECAUSE + e.getMessage());
            throw new CoreException("Uploaded file could not be imported, because: " + e.getMessage(), e);
        } finally {
            DataImporterUtils.deleteTempFile(file);
        }
        return importedList;
    }
	// END EXTRA CODE
}
