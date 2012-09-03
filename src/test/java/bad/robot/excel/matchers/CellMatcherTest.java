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

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hamcrest.StringDescription;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.firstRowOf;
import static bad.robot.excel.matchers.CellMatcher.hasSameCellAs;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.hamcrest.Matchers.not;

public class CellMatcherTest {

    private StringDescription description = new StringDescription();

    private Row row;

    @Before
    public void loadWorkbookAndSheets() throws IOException {
        row = firstRowOf("rowWithThreeCells.xls");
    }

    @Test
    public void exampleUsage() {
        assertThat(row, hasSameCellAs(createCell(0, 0, "C1, R1")));
        assertThat(row, not(hasSameCellAs(createCell(0, 0, "XXX"))));
    }

    @Test
    public void matches() {
        assertThat(hasSameCellAs(createCell(0, 1, "C2, R1")).matches(row), is(true));
    }

    @Test
    public void doesNotMatch() {
        assertThat(hasSameCellAs(createCell(0, 0, "C3, R1")).matches(row), is(false));
    }

    @Test
    public void describe() {
        hasSameCellAs(createCell(0, 0, "XXX")).describeTo(description);
        assertThat(description.toString(), is("equality of cell \"A1\""));
    }

    @Test
    public void mismatch() {
        hasSameCellAs(createCell(0, 0, "XXX")).matchesSafely(row, description);
        assertThat(description.toString(), is("cell at \"A1\" contained <\"C1, R1\"> expected <\"XXX\">"));
    }

    public static Cell createCell(int row, int column, String value) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        HSSFCell cell = sheet.createRow(row).createCell(column, CELL_TYPE_STRING);
        cell.setCellValue(value);
        return cell;
    }
}
