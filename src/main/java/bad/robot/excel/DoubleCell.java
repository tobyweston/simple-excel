package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

public class DoubleCell extends Cell {

    private final Double value;

    public DoubleCell(Double value) {
        this(value, new NoStyle());
    }

    public DoubleCell(Double value, Style style) {
        super(style);
        this.value = value;
    }

    @Override
    public void addTo(Row row, CellIndex index, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(index.value(), CELL_TYPE_NUMERIC);
        this.getStyle().applyTo(cell, workbook);
        cell.setCellValue(value);
    }

}
