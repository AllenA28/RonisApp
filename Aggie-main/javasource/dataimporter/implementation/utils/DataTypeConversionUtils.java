package dataimporter.implementation.utils;

import com.mendix.systemwideinterfaces.core.meta.IMetaPrimitive;
import dataimporter.implementation.enums.FileType;
import dataimporter.implementation.model.CellData;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.math.BigDecimal;
import java.util.Date;

public class DataTypeConversionUtils {

    private static final String INCOMPATIBLE_DATA_CONVERSION = "Incompatible data conversion";
    private static final String UNABLE_TO_PARSE_SOURCE_VALUE = "Unable to parse an source value : ";
    private static final String IN_TO = " in to ";

    private DataTypeConversionUtils() {

    }

    public static Object getConvertedData(CellData cellData, FileType fileType, IMetaPrimitive.PrimitiveType targetType) {
        var data = fileType.equals(FileType.EXCEL) ? cellData.getFormattedData() : cellData.getRawData();

        switch (targetType) {
            case String:
                return convertToString(cellData, data, targetType);
            case Integer:
                return convertToInt(targetType, data);
            case Long:
                return convertToLong(targetType, data);
            case Decimal:
                return convertToDecimal(targetType, data);
            case Boolean:
                return convertToBoolean(targetType, data);
            case DateTime:
                return convertToDateTime(targetType, data);
            default:
                return new DataImporterRuntimeException("Data type conversion not supported.");
        }
    }

    public static Object convertToString(CellData cellData, Object data, IMetaPrimitive.PrimitiveType targetType) {
        try {
            if (data instanceof String) {
                return data;
            } else if (data instanceof Date) {
                var dataFormatter = new DataFormatter();
                return dataFormatter.formatRawCellContents((Double) cellData.getRawData(), cellData.getDataFormat(), cellData.getDisplayMask());
            } else {
                return String.valueOf(data);
            }
        } catch (Exception exception) {
            throw new DataImporterRuntimeException(UNABLE_TO_PARSE_SOURCE_VALUE + data + IN_TO + targetType, exception);
        }
    }

    private static Object convertToDateTime(IMetaPrimitive.PrimitiveType targetType, Object data) {
        try {
            if (data instanceof Date) {
                return data;
            } else {
                throw new DataImporterRuntimeException(INCOMPATIBLE_DATA_CONVERSION);
            }
        } catch (Exception e) {
            throw new DataImporterRuntimeException(UNABLE_TO_PARSE_SOURCE_VALUE + data.toString() + IN_TO + targetType, e);
        }
    }

    private static Object convertToBoolean(IMetaPrimitive.PrimitiveType targetType, Object data) {
        try {
            if (data instanceof Boolean) {
                return data;
            } else if (data instanceof Date || data instanceof Double) {
                throw new DataImporterRuntimeException(INCOMPATIBLE_DATA_CONVERSION);
            } else {
                if (data.toString().equalsIgnoreCase("true") || data.toString().equalsIgnoreCase("false")) {
                    return Boolean.valueOf(data.toString());
                } else {
                    throw new DataImporterRuntimeException(INCOMPATIBLE_DATA_CONVERSION);
                }
            }
        } catch (Exception e) {
            throw new DataImporterRuntimeException(UNABLE_TO_PARSE_SOURCE_VALUE + data.toString() + IN_TO + targetType, e);
        }
    }

    private static Object convertToDecimal(IMetaPrimitive.PrimitiveType targetType, Object data) {
        try {
            if (data instanceof Boolean) {
                return new BigDecimal(String.valueOf(((Boolean) data).booleanValue()));
            } else if (data instanceof Date) {
                throw new DataImporterRuntimeException(INCOMPATIBLE_DATA_CONVERSION);
            } else if (data instanceof Double) {
                return new BigDecimal(String.valueOf(((Double) data).doubleValue()));
            } else {
                return new BigDecimal(data.toString());
            }
        } catch (Exception e) {
            throw new DataImporterRuntimeException(UNABLE_TO_PARSE_SOURCE_VALUE + data.toString() + IN_TO + targetType, e);
        }
    }

    private static Object convertToLong(IMetaPrimitive.PrimitiveType targetType, Object data) {
        try {
            if (data instanceof Boolean) {
                return Long.valueOf(String.valueOf(((Boolean) data).booleanValue()));
            } else if (data instanceof Date) {
                throw new DataImporterRuntimeException(INCOMPATIBLE_DATA_CONVERSION);
            } else if (data instanceof Double) {
                return new BigDecimal(String.valueOf(data)).longValueExact();
            } else {
                return Long.parseLong(data.toString());
            }
        } catch (Exception e) {
            throw new DataImporterRuntimeException(UNABLE_TO_PARSE_SOURCE_VALUE + data.toString() + IN_TO + targetType, e);
        }
    }

    private static Object convertToInt(IMetaPrimitive.PrimitiveType targetType, Object data) {
        try {
            if (data instanceof Boolean) {
                return Integer.valueOf(String.valueOf(((Boolean) data).booleanValue()));
            } else if (data instanceof Date) {
                throw new DataImporterRuntimeException(INCOMPATIBLE_DATA_CONVERSION);
            } else if (data instanceof Double) {
                return new BigDecimal(String.valueOf(data)).intValueExact();
            } else {
                return Integer.parseInt(data.toString());
            }
        } catch (Exception e) {
            throw new DataImporterRuntimeException(UNABLE_TO_PARSE_SOURCE_VALUE + data.toString() + IN_TO + targetType, e);
        }
    }

}
