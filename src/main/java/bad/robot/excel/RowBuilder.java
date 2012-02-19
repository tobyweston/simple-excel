package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class RowBuilder {

    private Map<CellIndex,Cell> cells  = new HashMap<CellIndex,Cell>();

    private RowBuilder() {
    }

    public static RowBuilder aRow() {
        return new RowBuilder();
    }

    public RowBuilder withString(CellIndex index, String text) {
        this.cells.put(index, new StringCell(text));
        return this;
    }

    public RowBuilder withDouble(CellIndex index, Double value) {
        this.cells.put(index, new DoubleCell(value));
        return this;
    }

    public RowBuilder withDouble(CellIndex index, Double value, Style style) {
        this.cells.put(index, new DoubleCell(value, style));
        return this;
    }

    public RowBuilder withInteger(CellIndex index, Integer value) {
        this.cells.put(index, new DoubleCell(new Double(Integer.toString(value))));
        return this;
    }

    public RowBuilder withDate(CellIndex index, Date date) {
        this.cells.put(index, new DateCell(date));
        return this;
    }

    public RowBuilder withDate(CellIndex index, Date date, Style style) {
        this.cells.put(index, new DateCell(date, style));
        return this;
    }

    public Row build() {
        return new Row(cells);
    }
}
