package dataimporter.factory;

import com.mendix.core.CoreException;
import dataimporter.implementation.enums.FileExtension;
import dataimporter.implementation.utils.DataImporterUtils;

public class DataImporterFactory {
    private DataImporterFactory() {
    }

    public static IDataProcessor getDataProcessor(FileExtension fileType) throws CoreException {
        switch (fileType) {
            case XLS:
                return new ExcelDataProcessor();
            case XLSX:
                return new ExcelXlsxDataProcessor();
            case CSV:
                return new CsvDataProcessor();
            case UNKNOWN:
            default: {
                DataImporterUtils.logNode.error("File extension is not an Excel or CSV extension ('.xls' or '.xlsx' or '.csv').");
                throw new CoreException("File extension is not an Excel or CSV extension ('.xls' or '.xlsx' or '.csv').");
            }
        }
    }
}
