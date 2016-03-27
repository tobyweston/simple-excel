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
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;

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
import static org.hamcrest.Matchers.instanceOf;
import static org.hamcrest.Matchers.is;

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
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceCellTemplate.xls")).replaceCell(coordinate(column(getColumn("A")), row(1)), "Hello World");
        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceCellTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceDateCell() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceDateCellTemplate.xls")).replaceCell(coordinate(column(getColumn("A")), row(1)), createDate(22, MAY, 1997));
        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceDateCellTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceCellsInComplicatedExample() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls"))
                .replaceCell(coordinate(column(getColumn("C")), row(5)), "Very")
                .replaceCell(coordinate(column(getColumn("D")), row(11)), "Complicated")
                .replaceCell(coordinate(column(getColumn("G")), row(3)), "Example")
                .replaceCell(coordinate(column(getColumn("H")), row(7)), "Of")
                .replaceCell(coordinate(column(getColumn("J")), row(10)), "Templated")
                .replaceCell(coordinate(column(getColumn("M")), row(15)), "Spreadsheet");

        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceCellsInComplicatedAlternateSyntaxExample() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls"))
                .replaceCell(coordinate(getColumn("C"), 5), "Very")
                .replaceCell(coordinate(getColumn("D"), 11), "Complicated")
                .replaceCell(coordinate(getColumn("G"), 3), "Example")
                .replaceCell(coordinate(getColumn("H"), 7), "Of")
                .replaceCell(coordinate(getColumn("J"), 10), "Templated")
                .replaceCell(coordinate(getColumn("M"), 15), "Spreadsheet");

        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls")));
    }

    @Test
    public void shouldAppendRow() throws IOException, ParseException {
        RowBuilder row = aRow()
                .withString(column(getColumn("A")), "This")
                .withString(column(getColumn("C")), "Row")
                .withString(column(getColumn("D")), "was")
                .withString(column(getColumn("E")), "inserted")
                .withString(column(getColumn("H")), "here")
                .withString(column(getColumn("I")), "by")
                .withString(column(getColumn("J")), "some")
                .withString(column(getColumn("K")), "smart")
                .withString(column(getColumn("N")), "Programmer")
                .withInteger(column(getColumn("L")), 1)
                .withDate(column(getColumn("P")), createDate(11, FEBRUARY, 1979))
                .withDouble(column(getColumn("O")), new Double("0.123456789"));
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("shouldAppendRowTemplate.xls")).appendRowToFirstSheet(row.build());

        assertThat(modified.workbook(), sameWorkbook(getWorkbook("shouldAppendRowTemplateExpected.xls")));
    }

    @Test
    public void replaceCellWithSameCell() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("cellTypes.xls")).replaceCell(coordinate(getColumn("B"), 4), "value");
        assertThat(modified.workbook(), sameWorkbook(getWorkbook("cellTypes.xls")));
    }

    @Test
    public void replaceCellWithSameRow() throws IOException {
        RowBuilder row = aRow()
                .withString(column(getColumn("A")), "String")
                .withString(column(getColumn("B")), "value");

        PoiWorkbook modified = new PoiWorkbook(getWorkbook("cellTypes.xls")).appendRowToFirstSheet(row.build());
        assertThat(modified.workbook(), Matchers.sameWorkbook(getWorkbook("replaceCellWithSameRowExpected.xls")));
    }

    @Test
    public void replaceThenLoadBlankCell() throws IOException {
        PoiWorkbook modified = new PoiWorkbook(getWorkbook("cellTypes.xls")).blankCell(coordinate(getColumn("A"), 1));
        assertThat(adaptPoi(getCellForCoordinate(coordinate(getColumn("A"), 1), modified.workbook())), is(instanceOf(BlankCell.class)));
    }

}