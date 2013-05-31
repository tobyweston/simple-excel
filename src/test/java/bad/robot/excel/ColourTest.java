package bad.robot.excel;

import bad.robot.excel.valuetypes.Coordinate;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.Colour.Blue;
import static bad.robot.excel.Fill.fill;
import static bad.robot.excel.ForegroundColour.foregroundColour;
import static bad.robot.excel.StyleBuilder.aStyle;
import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.valuetypes.ExcelColumnIndex.A;

public class ColourTest {

    @Test
    @Ignore
    public void coloursAreApplied() throws IOException {
        Fill fill = fill(foregroundColour(Blue));
        Cell cell = new StringCell("This should be BLUE", aStyle().with(fill));
        Workbook workbook = getWorkbook("emptySheet.xlsx");
        new PoiWorkbookMutator(workbook).replaceCell(Coordinate.coordinate(A, 1), cell);
        OutputWorkbook.writeWorkbookToTemporaryFile(workbook, "test");
        // no assertions yet, hence the ignore
    }
}
