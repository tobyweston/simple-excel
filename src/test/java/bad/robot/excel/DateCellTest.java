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

package bad.robot.excel;

import bad.robot.excel.cell.DateCell;
import bad.robot.excel.cell.DoubleCell;
import bad.robot.excel.sheet.Coordinate;
import bad.robot.excel.workbook.PoiWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.DateUtil.createDate;
import static bad.robot.excel.WorkbookResource.*;
import static bad.robot.excel.column.ColumnIndex.column;
import static bad.robot.excel.column.ExcelColumnIndex.*;
import static bad.robot.excel.matchers.CellMatcher.equalTo;
import static bad.robot.excel.sheet.Coordinate.coordinate;
import static java.util.Calendar.MARCH;
import static java.util.Calendar.OCTOBER;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;

public class DateCellTest {

    private final DateCell cell = new DateCell(createDate(12, OCTOBER, 2013));

    @Test
    public void shouldSetDataFormatWhenAddingACell() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(0);
        cell.addTo(row, column("A"), workbook);
        assertThat(getCellDataFormatAtCoordinate(coordinate("A", 1), workbook), is("dd-MMM-yyyy"));
    }

    @Test
    public void shouldSetDataFormatWhenReplacingACell() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(0);
        HSSFCell original = row.createCell(1);
        cell.update(original, workbook);
        assertThat(getCellDataFormatAtCoordinate(coordinate("B", 1), workbook), is("dd-MMM-yyyy"));
    }

    @Test
    public void replaceNonDateCellWithACell() throws IOException {
        Workbook workbook = getWorkbook("cellTypes.xls");
        PoiWorkbook sheet = new PoiWorkbook(workbook);
        Coordinate coordinate = coordinate("B", 2);

        assertThat(getCellForCoordinate(coordinate, workbook), equalTo(new DoubleCell(1001d)));
        sheet.replaceCell(coordinate, createDate(15, MARCH, 2012));

        assertThat(getCellForCoordinate(coordinate, workbook), equalTo(new DateCell(createDate(15, MARCH, 2012))));
        assertThat(getCellDataFormatAtCoordinate(coordinate, workbook), is("dd-MMM-yyyy"));
        assertThat("should not have affected a shared data format", getCellDataFormatAtCoordinate(coordinate("B", 7), workbook), is("General"));
    }
}
