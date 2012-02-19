package bad.robot.excel;

import bad.robot.excel.valuetypes.Coordinate;
import bad.robot.excel.valuetypes.RowIndex;
import bad.robot.excel.valuetypes.SheetIndex;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

import static org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK;


public class PoiWorkbookMutator implements WorkbookMutator {

    private final Workbook workbook;

    public PoiWorkbookMutator(Workbook workbook) {
        if (workbook == null)
            throw new IllegalArgumentException();
        this.workbook = workbook;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, String value) {
        getCellForCoordinate(coordinate).setCellValue(value);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Formula formula) {
        getCellForCoordinate(coordinate).setCellFormula(formula.value());
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Date value) {
        getCellForCoordinate(coordinate).setCellValue(value);
        return this;
    }

    @Override
    public WorkbookMutator replaceCell(Coordinate coordinate, Double value) {
        getCellForCoordinate(coordinate).setCellValue(value.doubleValue());
        return this;
    }

    @Override
    public WorkbookMutator copyRow(Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to) {
        CopyRow.copyRow(workbook, worksheet, from, to);
        return this;
    }

    @Override
    public WorkbookMutator insertRowToFirstSheet(Row row, RowIndex index) {
        row.insertAt(workbook, SheetIndex.sheet(0), index);
        return this;
    }

    @Override
    public WorkbookMutator insertRowToSheet(Row row, RowIndex index, SheetIndex sheet) {
        row.insertAt(workbook, sheet, index);
        return this;
    }

    @Override
    public WorkbookMutator appendRowToFirstSheet(Row row) {
        row.appendTo(workbook, SheetIndex.sheet(0));
        return this;
    }

    @Override
    public WorkbookMutator appendRowToSheet(Row row, SheetIndex index) {
        row.appendTo(workbook, index);
        return this;
    }

    @Override
    public WorkbookMutator refreshFormulas() {
        HSSFFormulaEvaluator.evaluateAllFormulaCells((HSSFWorkbook) workbook);
        return this;
    }

    private Cell getCellForCoordinate(Coordinate coordinate) {
        Sheet sheet = workbook.getSheetAt(coordinate.getSheet().value());
        org.apache.poi.ss.usermodel.Row row = sheet.getRow(coordinate.getRow().value());
        if (row == null)
            row = sheet.createRow(coordinate.getRow().value());
        return row.getCell(coordinate.getColumn().value(), CREATE_NULL_AS_BLANK);
    }
}