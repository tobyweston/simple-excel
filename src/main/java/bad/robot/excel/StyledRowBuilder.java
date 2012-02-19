package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class StyledRowBuilder {

    private Map<CellIndex,Cell> cells  = new HashMap<CellIndex,Cell>();
    private Style defaultStyle;

    public static StyledRowBuilder aRowWithDefaultStyle(Style defaultStyle) {
        return new StyledRowBuilder(defaultStyle);
    }

    public static StyledRowBuilder aRow() {
        return new StyledRowBuilder(new NoStyle());
    }

    private StyledRowBuilder(Style defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public StyledRowBuilder withBlank(CellIndex index) {
        this.cells.put(index, new BlankCell(defaultStyle));
        return this;
    }

    public StyledRowBuilder withBlank(CellIndex index, Style style) {
        this.cells.put(index, new BlankCell(style));
        return this;
    }

    public StyledRowBuilder withString(CellIndex index, String text) {
        this.cells.put(index, new StringCell(text, defaultStyle));
        return this;
    }

    public StyledRowBuilder withString(CellIndex index, String text, Style style) {
        this.cells.put(index, new StringCell(text, style));
        return this;
    }

    public StyledRowBuilder withDouble(CellIndex index, Double value) {
        this.cells.put(index, new DoubleCell(value, defaultStyle));
        return this;
    }

    public StyledRowBuilder withDouble(CellIndex index, Double value, Style style) {
        this.cells.put(index, new DoubleCell(value, style));
        return this;
    }

    public StyledRowBuilder withInteger(CellIndex index, Integer value) {
        this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), defaultStyle));
        return this;
    }

    public StyledRowBuilder withInteger(CellIndex index, Integer value, Style style) {
        this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), style));
        return this;
    }

    public StyledRowBuilder withDate(CellIndex index, Date date) {
        this.cells.put(index, new DateCell(date, defaultStyle));
        return this;
    }

    public StyledRowBuilder withDate(CellIndex index, Date date, Style style) {
        this.cells.put(index, new DateCell(date, style));
        return this;
    }

    public StyledRowBuilder withFormula(CellIndex index, String formula) {
        this.cells.put(index, new FormulaCell(formula, defaultStyle));
        return this;
    }

    public StyledRowBuilder withFormula(CellIndex index, String formula, Style style) {
        this.cells.put(index, new FormulaCell(formula, style));
        return this;
    }


    public Row build() {
        return new Row(cells);
    }
}
