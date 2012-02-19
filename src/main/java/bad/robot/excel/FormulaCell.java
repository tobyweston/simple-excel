package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;

public class FormulaCell extends Cell {

    private final String formula;

    public FormulaCell(String formula) {
        this(formula, new NoStyle());
    }

    public FormulaCell(String text, Style style) {
        super(style);
        this.formula = text;
    }

    @Override
    public void addTo(Row row, CellIndex index, Workbook workbook) {
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(index.value(), CELL_TYPE_FORMULA);
        this.getStyle().applyTo(cell, workbook);
        cell.setCellFormula(formula);
    }

}
