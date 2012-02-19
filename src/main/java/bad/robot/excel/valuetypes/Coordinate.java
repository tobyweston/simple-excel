package bad.robot.excel.valuetypes;

import static bad.robot.excel.valuetypes.SheetIndex.sheet;

public class Coordinate {

    private final ColumnIndex column;
    private final RowIndex row;
    private final SheetIndex sheet;

    public static Coordinate coordinate(ColumnIndex column, RowIndex row, SheetIndex sheet) {
        return new Coordinate(column, row, sheet);
    }

    public static Coordinate coordinate(ColumnIndex column, RowIndex row) {
        return new Coordinate(column, row, sheet(0));
    }

    private Coordinate(ColumnIndex column, RowIndex row, SheetIndex sheet) {
        this.column = column;
        this.row = row;
        this.sheet = sheet;
    }

    public ColumnIndex getColumn() {
        return column;
    }

    public RowIndex getRow() {
        return row;
    }

    public SheetIndex getSheet() {
        return sheet;
    }
}