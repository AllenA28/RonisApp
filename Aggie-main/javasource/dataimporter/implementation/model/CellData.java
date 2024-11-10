package dataimporter.implementation.model;

import java.util.Objects;

public class CellData {
    private final int columnIndex;
    private final String columnHeader;
    private final Object rawData;
    private final String displayMask;
    private final Object formattedData;
    private final short dataFormat;

    public CellData(int columnIndex, String columnHeader, Object rawData, Object formattedData) {
        this(columnIndex, columnHeader, rawData, formattedData, null, (short) 0);
    }

    public CellData(int columnIndex, String columnHeader, Object rawData, Object formattedData, String displayMask, short dataFormat) {
        this.columnIndex = columnIndex;
        this.columnHeader = columnHeader;
        this.rawData = rawData;
        this.formattedData = formattedData;
        this.displayMask = displayMask;
        this.dataFormat = dataFormat;
    }

    public CellData(CellData cellData) {
        this.columnIndex = cellData.getColumnIndex();
        this.columnHeader = cellData.getColumnHeader();
        this.rawData = cellData.getRawData();
        this.formattedData = cellData.getFormattedData();
        this.displayMask = cellData.getDisplayMask();
        this.dataFormat = cellData.getDataFormat();
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public String getColumnHeader() {
        return columnHeader;
    }

    public Object getRawData() {
        return rawData;
    }

    public String getDisplayMask() {
        return displayMask;
    }

    public Object getFormattedData() {
        return formattedData;
    }

    public short getDataFormat() {
        return dataFormat;
    }


    @Override
    public int hashCode() {
        return columnIndex + 31 * Objects.hash(rawData, displayMask, formattedData);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CellData that = (CellData) o;
        return columnIndex == that.columnIndex &&
                columnHeader == that.columnHeader &&
                Objects.equals(rawData, that.rawData) &&
                Objects.equals(displayMask, that.displayMask) &&
                Objects.equals(formattedData, that.formattedData) &&
                Objects.equals(dataFormat, this.dataFormat);
    }

    @Override
    public String toString() {
        return "CellData{ " +
                "colNo=" + columnIndex +
                ", colName=" + columnHeader +
                ", rawData=" + rawData +
                ", formattedData=" + formattedData +
                ", displayMask='" + displayMask + '\'' +
                ", dataFormat=" + dataFormat +
                " }";
    }
}