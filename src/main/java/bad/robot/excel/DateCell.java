package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

public class DateCell extends Cell {

    private final Date value;

    public DateCell(Date date) {
        this(date, new NoStyle());
    }

    public DateCell(Date date, Style style) {
        super(style);
        this.value = date;
    }

    @Override
    public void addTo(Row row, CellIndex index, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(index.value()); // what poi type is the date cell?
        this.getStyle().applyTo(cell, workbook);
        cell.setCellValue(value);
    }

}
