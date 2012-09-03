/*
 * Copyright (c) 2012, bad robot (london) ltd.
 *
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package bad.robot.excel.matchers;

import org.apache.poi.ss.usermodel.Row;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.DateUtil.createDate;
import static bad.robot.excel.WorkbookResource.firstRowOf;
import static bad.robot.excel.matchers.CellMatcher.hasSameCellAs;
import static bad.robot.excel.matchers.StubCell.*;
import static java.util.Calendar.AUGUST;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class CellMatcherTest {

    private final StringDescription description = new StringDescription();

    private Row row;

    @Before
    public void loadWorkbookAndSheets() throws IOException {
        row = firstRowOf("rowWithVariousCells.xls");
    }

    @Test
    public void exampleUsage() {
        assertThat(row, hasSameCellAs(createCell(0, 6, "Text")));
        assertThat(row, not(hasSameCellAs(createCell(0, 6, "XXX"))));
    }

    @Test
    public void matches() {
        assertThat(hasSameCellAs(createBlankCell(0, 0)).matches(row), is(true));
        assertThat(hasSameCellAs(createCell(0, 1, true)).matches(row), is(true));
        assertThat(hasSameCellAs(createCell(0, 2, (byte) 0x07)).matches(row), is(true));
        assertThat(hasSameCellAs(createFormulaCell(0, 3, "2+3")).matches(row), is(true));
        assertThat(hasSameCellAs(createCell(0, 4, 34.5D)).matches(row), is(true));
        assertThat(hasSameCellAs(createCell(0, 5, createDate(22, AUGUST, 2012))).matches(row), is(true));
        assertThat(hasSameCellAs(createCell(0, 6, "Text")).matches(row), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameCellAs(createBlankCell(0, 1)).matches(row), is(false));
        assertThat(hasSameCellAs(createCell(0, 1, false)).matches(row), is(false));
        assertThat(hasSameCellAs(createCell(0, 3, (byte) 0x07)).matches(row), is(false));
        assertThat(hasSameCellAs(createFormulaCell(0, 3, "21*3")).matches(row), is(false));
        assertThat(hasSameCellAs(createCell(0, 4, 342.5D)).matches(row), is(false));
        assertThat(hasSameCellAs(createCell(0, 5, createDate(22, AUGUST, 2011))).matches(row), is(false));
        assertThat(hasSameCellAs(createCell(0, 6, "text")).matches(row), is(false));
    }

    @Test
    public void describe() {
        hasSameCellAs(createCell(0, 0, "XXX")).describeTo(description);
        assertThat(description.toString(), is("equality of cell \"A1\""));
    }

    @Test
    public void mismatchMissingCell() {
        hasSameCellAs(createCell(0, 0, "XXX")).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"A1\" contained <nothing> expected <\"XXX\">"));
    }

    @Test
    public void mismatchBooleanCell() {
        hasSameCellAs(createCell(0, 1, false)).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"B1\" contained <TRUE> expected <FALSE>"));
    }

    @Test
    public void mismatchFormulaErrorCell() {
        hasSameCellAs(createCell(0, 2, (byte) 0x00)).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"C1\" contained <Error:7> expected <Error:0>"));
    }

    @Test
    public void mismatchFormulaCell() {
        hasSameCellAs(createFormulaCell(0, 3, "2+2")).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"D1\" contained <Formula:2+3> expected <Formula:2+2>"));
    }

    @Test
    public void mismatchNumericCell() {
        hasSameCellAs(createCell(0, 4, 341.5D)).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"E1\" contained <34.5D> expected <341.5D>"));
    }

    @Test
    public void mismatchDateCell() {
        hasSameCellAs(createCell(0, 5, createDate(21, AUGUST, 2012))).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"F1\" contained <Wed Aug 22 00:00:00 BST 2012> expected <Tue Aug 21 00:00:00 BST 2012>"));
    }

    @Test
    public void mismatchStringCell() {
        hasSameCellAs(createCell(0, 6, "XXX")).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"G1\" contained <\"Text\"> expected <\"XXX\">"));
    }

}
