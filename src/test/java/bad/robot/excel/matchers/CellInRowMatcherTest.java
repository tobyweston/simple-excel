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

package bad.robot.excel.matchers;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.DateUtil.createDate;
import static bad.robot.excel.WorkbookResource.firstRowOf;
import static bad.robot.excel.matchers.Assertions.assertTimezone;
import static bad.robot.excel.matchers.CellInRowMatcher.hasSameCell;
import static bad.robot.excel.matchers.StubCell.*;
import static java.util.Calendar.AUGUST;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class CellInRowMatcherTest {

    private final StringDescription description = new StringDescription();

    private Sheet sheet;
    private Row row;


    @Before
    public void loadWorkbookAndSheets() throws IOException {
        row = firstRowOf("rowWithVariousCells.xls");
        sheet = row.getSheet();
    }

    @Test
    public void exampleUsage() {
        assertThat(row, hasSameCell(sheet, createCell(0, 6, "Text")));
        assertThat(row, not(hasSameCell(sheet, createCell(0, 6, "XXX"))));
    }

    @Test
    public void matches() {
        assertThat(hasSameCell(sheet, createBlankCell(0, 0)).matches(row), is(true));
        assertThat(hasSameCell(sheet, createCell(0, 1, true)).matches(row), is(true));
        assertThat(hasSameCell(sheet, createCell(0, 2, (byte) 0x07)).matches(row), is(true));
        assertThat(hasSameCell(sheet, createFormulaCell(0, 3, "2+3")).matches(row), is(true));
        assertThat(hasSameCell(sheet, createCell(0, 4, 34.5D)).matches(row), is(true));
        assertThat(hasSameCell(sheet, createCell(0, 5, createDate(22, AUGUST, 2012))).matches(row), is(true));
        assertThat(hasSameCell(sheet, createCell(0, 6, "Text")).matches(row), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameCell(sheet, createBlankCell(0, 1)).matches(row), is(false));
        assertThat(hasSameCell(sheet, createCell(0, 1, false)).matches(row), is(false));
        assertThat(hasSameCell(sheet, createCell(0, 3, (byte) 0x07)).matches(row), is(false));
        assertThat(hasSameCell(sheet, createFormulaCell(0, 3, "21*3")).matches(row), is(false));
        assertThat(hasSameCell(sheet, createCell(0, 4, 342.5D)).matches(row), is(false));
        assertThat(hasSameCell(sheet, createCell(0, 5, createDate(22, AUGUST, 2011))).matches(row), is(false));
        assertThat(hasSameCell(sheet, createCell(0, 6, "text")).matches(row), is(false));
    }

    @Test
    public void describe() {
        hasSameCell(sheet, createCell(0, 0, "XXX")).describeTo(description);
        assertThat(description.toString(), is("equality of cell \"A1\""));
    }

    @Test
    public void mismatchMissingCell() {
        hasSameCell(sheet, createCell(0, 0, "XXX")).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"A1\" contained <nothing> expected <\"XXX\"> sheet \"Sheet1\""));
    }

    @Test
    public void mismatchBooleanCell() {
        hasSameCell(sheet, createCell(0, 1, false)).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"B1\" contained <TRUE> expected <FALSE> sheet \"Sheet1\""));
    }

    @Test
    public void mismatchFormulaErrorCell() {
        hasSameCell(sheet, createCell(0, 2, (byte) 0x00)).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"C1\" contained <Error:7> expected <Error:0> sheet \"Sheet1\""));
    }

    @Test
    public void mismatchFormulaCell() {
        hasSameCell(sheet, createFormulaCell(0, 3, "2+2")).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"D1\" contained <Formula:2+3> expected <Formula:2+2> sheet \"Sheet1\""));
    }

    @Test
    public void mismatchNumericCell() {
        hasSameCell(sheet, createCell(0, 4, 341.5D)).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"E1\" contained <34.5D> expected <341.5D> sheet \"Sheet1\""));
    }

    @Test
    public void mismatchDateCell() {
        assertTimezone(is("UTC"));
        hasSameCell(sheet, createCell(0, 5, createDate(21, AUGUST, 2012))).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"F1\" contained <Wed Aug 22 00:00:00 UTC 2012> expected <Tue Aug 21 00:00:00 UTC 2012> sheet \"Sheet1\""));
    }

    @Test
    public void mismatchStringCell() {
        hasSameCell(sheet, createCell(0, 6, "XXX")).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"G1\" contained <\"Text\"> expected <\"XXX\"> sheet \"Sheet1\""));
    }

}
