/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package bad.robot.excel.workbook;

import bad.robot.excel.WorkbookResource;
import bad.robot.excel.cell.BlankCell;
import bad.robot.excel.matchers.Matchers;
import bad.robot.excel.row.RowBuilder;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;

import static bad.robot.excel.DateUtil.createDate;
import static bad.robot.excel.WorkbookResource.getCellForCoordinate;
import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.column.ColumnIndex.column;
import static bad.robot.excel.column.ExcelColumnIndex.*;
import static bad.robot.excel.matchers.CellType.adaptPoi;
import static bad.robot.excel.matchers.WorkbookMatcher.sameWorkbook;
import static bad.robot.excel.row.DefaultRowBuilder.aRow;
import static bad.robot.excel.row.RowIndex.row;
import static bad.robot.excel.sheet.Coordinate.coordinate;
import static bad.robot.excel.workbook.WorkbookType.XLS;
import static bad.robot.excel.workbook.WorkbookType.XML;
import static java.util.Calendar.FEBRUARY;
import static java.util.Calendar.MAY;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.*;

public class PoiWorkbookTest {

    @Test(expected = IllegalArgumentException.class)
    public void shouldNotAllowNullWorkbook() {
        new PoiWorkbook((Workbook) null);
    }

    @Test
    public void loadWorkbookFromStream() throws IOException {
        InputStream stream = WorkbookResource.class.getResourceAsStream("emptySheet.xls");
        new PoiWorkbook(stream);
    }

    @Test
    public void createXlsSheet() throws IOException {
        PoiWorkbook workbook = new PoiWorkbook(XLS);
        assertThat(workbook.workbook(), sameWorkbook(getWorkbook("emptySheet.xls")));
    }

    @Test
    public void createXmlSheet() throws IOException {
        PoiWorkbook workbook = new PoiWorkbook(XML);
        assertThat(workbook.workbook(), sameWorkbook(getWorkbook("emptySheet.xlsx")));
    }

    @Test
    public void shouldReplaceCell() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceCellTemplate.xls")).replaceCell(coordinate(column(A), row(1)), "Hello World");
        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceCellTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceDateCell() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceDateCellTemplate.xls")).replaceCell(coordinate(column(A), row(1)), createDate(22, MAY, 1997));
        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceDateCellTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceCellsInComplicatedExample() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls"))
                .replaceCell(coordinate(column(C), row(5)), "Very")
                .replaceCell(coordinate(column(D), row(11)), "Complicated")
                .replaceCell(coordinate(column(G), row(3)), "Example")
                .replaceCell(coordinate(column(H), row(7)), "Of")
                .replaceCell(coordinate(column(J), row(10)), "Templated")
                .replaceCell(coordinate(column(M), row(15)), "Spreadsheet");

        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceCellsInComplicatedAlternateSyntaxExample() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls"))
                .replaceCell(coordinate(C, 5), "Very")
                .replaceCell(coordinate(D, 11), "Complicated")
                .replaceCell(coordinate(G, 3), "Example")
                .replaceCell(coordinate(H, 7), "Of")
                .replaceCell(coordinate(J, 10), "Templated")
                .replaceCell(coordinate(M, 15), "Spreadsheet");

        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls")));
    }

    @Test
    public void shouldAppendRow() throws IOException {
        RowBuilder row = aRow()
                .withString(column(A), "This")
                .withString(column(C), "Row")
                .withString(column(D), "was")
                .withString(column(E), "inserted")
                .withString(column(H), "here")
                .withString(column(I), "by")
                .withString(column(J), "some")
                .withString(column(K), "smart")
                .withString(column(N), "Programmer")
                .withInteger(column(L), 1)
                .withDate(column(P), createDate(11, FEBRUARY, 1979))
                .withDouble(column(O), new Double("0.123456789"));
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldAppendRowTemplate.xls")).appendRowToFirstSheet(row.build());

        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldAppendRowTemplateExpected.xls")));
    }

    @Test
    public void replaceCellWithSameCell() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("cellTypes.xls")).replaceCell(coordinate(B, 4), "value");
        assertThat(modified.workbook(), sameWorkbook(getWorkbook("cellTypes.xls")));
    }

    @Test
    public void replaceCellWithSameRow() throws IOException {
        RowBuilder row = aRow()
                .withString(column(A), "String")
                .withString(column(B), "value");

        PoiWorkbook modified = new PoiWorkbook(getWorkbook("cellTypes.xls")).appendRowToFirstSheet(row.build());
        assertThat(modified.workbook(), Matchers.sameWorkbook(getWorkbook("replaceCellWithSameRowExpected.xls")));
    }

    @Test
    public void replaceThenLoadBlankCell() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("cellTypes.xls")).blankCell(coordinate(A, 1));
        assertThat(adaptPoi(getCellForCoordinate(coordinate(A, 1), modified.workbook())), is(instanceOf(BlankCell.class)));
    }

    @Test
    @Ignore
    public void shouldCountMergedCellsWhenSingleEmptySheetHasMergedCells() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        int firstMergedCol = 1;
        int lastMergedCol = 5;
        sheet.addMergedRegion(new CellRangeAddress(0, 0, firstMergedCol, lastMergedCol));

        assertThat(workbook, is(not(sameWorkbook(getWorkbook("emptySheet.xlsx")))));
    }
}
