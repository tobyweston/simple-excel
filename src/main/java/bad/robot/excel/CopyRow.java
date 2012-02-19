package bad.robot.excel;

import bad.robot.excel.valuetypes.RowIndex;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

public class CopyRow {

    public static void copyRow(Workbook workbook, Sheet worksheet, RowIndex from, RowIndex to) {
        Row sourceRow = worksheet.getRow(from.value());
        Row newRow = worksheet.getRow(to.value());

        if (alreadyExists(newRow))
            worksheet.shiftRows(to.value(), worksheet.getLastRowNum(), 1);
        else
            newRow = worksheet.createRow(to.value());

        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);
            if (oldCell != null) {
                copyCellStyle(workbook, oldCell, newCell);
                copyCellComment(oldCell, newCell);
                copyCellHyperlink(oldCell, newCell);
                copyCellDataTypeAndValue(oldCell, newCell);
            }
        }

        copyAnyMergedRegions(worksheet, sourceRow, newRow);
    }

    private static void copyCellStyle(Workbook workbook, Cell oldCell, Cell newCell) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(oldCell.getCellStyle());
        newCell.setCellStyle(style);
    }

    private static void copyCellComment(Cell oldCell, Cell newCell) {
        if (newCell.getCellComment() != null)
            newCell.setCellComment(oldCell.getCellComment());
    }

    private static void copyCellHyperlink(Cell oldCell, Cell newCell) {
        if (oldCell.getHyperlink() != null)
            newCell.setHyperlink(oldCell.getHyperlink());
    }

    private static void copyCellDataTypeAndValue(Cell oldCell, Cell newCell) {
        setCellDataType(oldCell, newCell);
        setCellDataValue(oldCell, newCell);
    }

    private static void setCellDataType(Cell oldCell, Cell newCell) {
        newCell.setCellType(oldCell.getCellType());
    }

    private static void setCellDataValue(Cell oldCell, Cell newCell) {
        switch (oldCell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING:
                newCell.setCellValue(oldCell.getRichStringCellValue());
                break;
        }
    }

    private static boolean alreadyExists(Row newRow) {
        return newRow != null;
    }

    private static void copyAnyMergedRegions(Sheet worksheet, Row sourceRow, Row newRow) {
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++)
            copyMergeRegion(worksheet, sourceRow, newRow, worksheet.getMergedRegion(i));
    }

    private static void copyMergeRegion(Sheet worksheet, Row sourceRow, Row newRow, CellRangeAddress mergedRegion) {
        CellRangeAddress range = mergedRegion;
        if (range.getFirstRow() == sourceRow.getRowNum()) {
            int lastRow = newRow.getRowNum() + (range.getFirstRow() - range.getLastRow());
            CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(), lastRow, range.getFirstColumn(), range.getLastColumn());
            worksheet.addMergedRegion(newCellRangeAddress);
        }
    }

}