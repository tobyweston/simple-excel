package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK;

public class BlankCell extends Cell {

    public BlankCell() {
        this(new NoStyle());
    }

    public BlankCell(Style style) {
        super(style);
    }

    @Override
    public void addTo(Row row, CellIndex index, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(index.value(), CELL_TYPE_BLANK);
        this.getStyle().applyTo(cell, workbook);
    }

}
