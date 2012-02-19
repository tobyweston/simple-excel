package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class Cell {

    private final Style style;

    protected Cell(Style style) {
        this.style = style;
    }

    public Style getStyle() {
        return style;
    }

    public abstract void addTo(Row row, CellIndex index, Workbook workbook);

}
