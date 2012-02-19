package bad.robot.excel;

import bad.robot.excel.valuetypes.CellIndex;
import bad.robot.excel.valuetypes.RowIndex;
import bad.robot.excel.valuetypes.SheetIndex;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class Row {

    private static final int shiftDownAmount = 1;

    private final Map<CellIndex, Cell> cells = new HashMap<CellIndex, Cell>();

    public Row(Map<CellIndex, Cell> cells) {
        this.cells.putAll(cells);
    }

    public void insertAt(Workbook workbook, SheetIndex sheetIndex, RowIndex rowIndex) {
        Sheet sheet = workbook.getSheetAt(sheetIndex.value());
        sheet.shiftRows(rowIndex.value(), sheet.getLastRowNum(), shiftDownAmount);
        org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowIndex.value());
        copyCellsTo(row, workbook);
    }

    public void appendTo(Workbook workbook, SheetIndex index) {
        Sheet sheet = workbook.getSheetAt(index.value());
        org.apache.poi.ss.usermodel.Row row = createRow(sheet);
        copyCellsTo(row, workbook);
    }

    private void copyCellsTo(org.apache.poi.ss.usermodel.Row row, Workbook workbook) {
        for (CellIndex index : cells.keySet()) {
            Cell cellToInsert = cells.get(index);
            cellToInsert.addTo(row, index, workbook);
        }
    }

    private static org.apache.poi.ss.usermodel.Row createRow(Sheet sheet) {
        if (sheet.getPhysicalNumberOfRows() == 0)
            return sheet.createRow(0);
        return sheet.createRow(sheet.getLastRowNum() + 1);
    }


}
