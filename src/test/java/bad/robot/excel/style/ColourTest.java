package bad.robot.excel.style;

import bad.robot.excel.OutputWorkbook;
import bad.robot.excel.cell.Cell;
import bad.robot.excel.cell.StringCell;
import bad.robot.excel.workbook.PoiWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.column.ExcelColumnIndex.A;
import static bad.robot.excel.sheet.Coordinate.coordinate;
import static bad.robot.excel.style.Colour.*;
import static bad.robot.excel.style.Fill.fill;
import static bad.robot.excel.style.FontColour.fontColour;
import static bad.robot.excel.style.FontSize.fontSize;
import static bad.robot.excel.style.ForegroundColour.foregroundColour;
import static bad.robot.excel.style.StyleBuilder.aStyle;

public class ColourTest {

    @Test
    public void exampleUsage() throws IOException {
        StyleBuilder blue = aStyle().with(fill(foregroundColour(Blue))).with(fontSize("18")).with(fontColour(White));
        StyleBuilder yellow = aStyle().with(fill(foregroundColour(Yellow))).with(fontColour(White));
        Cell a1 = new StringCell("Blue", blue);
        Cell a2 = new StringCell("Yellow", yellow);
        Workbook workbook = getWorkbook("emptySheet.xlsx");
        PoiWorkbook sheet = new PoiWorkbook(workbook);
        sheet.replaceCell(coordinate(A, 1), a1);
        sheet.replaceCell(coordinate(A, 2), a2);
        OutputWorkbook.writeWorkbookToTemporaryFile(workbook, "exampleOfApplyColours");
    }
}
