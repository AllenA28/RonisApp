package dataimporter.implementation.utils;

public class DataImporterRuntimeException extends RuntimeException {

    public DataImporterRuntimeException(String message) {
        super(message);
    }

    public DataImporterRuntimeException(String message, Exception exception) {
        super(message, exception);
    }
}

