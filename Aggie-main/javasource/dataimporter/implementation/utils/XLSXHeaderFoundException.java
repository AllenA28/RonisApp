package dataimporter.implementation.utils;

import org.xml.sax.SAXException;

public class XLSXHeaderFoundException extends SAXException {
    public XLSXHeaderFoundException(String message) {
        super(message);
    }
}
