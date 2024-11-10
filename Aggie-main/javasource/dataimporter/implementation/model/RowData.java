package dataimporter.implementation.model;

import java.util.List;

public class RowData {
    long rowIndex;
    List<CellData> cellData;

    public long getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(long rowIndex) {
        this.rowIndex = rowIndex;
    }

    public List<CellData> getCellData() {
        return cellData;
    }

    public void setCellData(List<CellData> cellData) {
        this.cellData = cellData;
    }
}
