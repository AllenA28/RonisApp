package dataimporter.implementation.utils;

public class DataReaderException extends Exception {

    public DataReaderException(String message) {
        super(message);
    }

    public DataReaderException(String message, Exception exception) {
        super(message, exception);
    }
}
