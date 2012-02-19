package bad.robot.excel;

import bad.robot.excel.valuetypes.Coordinate;
import bad.robot.excel.valuetypes.RowIndex;
import bad.robot.excel.valuetypes.SheetIndex;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Date;

public interface WorkbookMutator {

    WorkbookMutator replaceCell(Coordinate coordinate, String value);

    WorkbookMutator replaceCell(Coordinate coordinate, Formula formula);

    WorkbookMutator replaceCell(Coordinate coordinate, Date value);

    WorkbookMutator replaceCell(Coordinate coordinate, Double value);

    WorkbookMutator copyRow(Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to);

    WorkbookMutator insertRowToFirstSheet(Row row, RowIndex index);

    WorkbookMutator insertRowToSheet(Row row, RowIndex index, SheetIndex sheet);

    WorkbookMutator appendRowToFirstSheet(Row row);

    WorkbookMutator appendRowToSheet(Row row, SheetIndex index);

    WorkbookMutator refreshFormulas();

}

