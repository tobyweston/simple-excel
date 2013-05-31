package bad.robot.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.Colour.*;
import static bad.robot.excel.Fill.fill;
import static bad.robot.excel.ForegroundColour.foregroundColour;
import static bad.robot.excel.StyleBuilder.aStyle;
import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.valuetypes.Coordinate.coordinate;
import static bad.robot.excel.valuetypes.ExcelColumnIndex.A;
import static bad.robot.excel.valuetypes.FontColour.fontColour;
import static bad.robot.excel.valuetypes.FontSize.fontSize;

public class ColourTest {

    @Test
    public void exampleUsage() throws IOException {
        StyleBuilder blue = aStyle().with(fill(foregroundColour(Blue))).with(fontSize("18")).with(fontColour(White));
        StyleBuilder yellow = aStyle().with(fill(foregroundColour(Yellow))).with(fontColour(White));
        Cell a1 = new StringCell("Blue", blue);
        Cell a2 = new StringCell("Yellow", yellow);
        Workbook workbook = getWorkbook("emptySheet.xlsx");
        PoiWorkbookMutator sheet = new PoiWorkbookMutator(workbook);
        sheet.replaceCell(coordinate(A, 1), a1);
        sheet.replaceCell(coordinate(A, 2), a2);
        OutputWorkbook.writeWorkbookToTemporaryFile(workbook, "exampleOfApplyColours");
    }
}
