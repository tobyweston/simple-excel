package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;

public class StringCell extends Cell {

    private final String value;

    public StringCell(String text) {
        this(text, new NoStyle());
    }

    public StringCell(String text, Style style) {
        super(style);
        this.value = text;
    }

    @Override
    public void addTo(Row row, CellIndex index, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(index.value(), CELL_TYPE_STRING);
        this.getStyle().applyTo(cell, workbook);
        cell.setCellValue(value);
    }

}
