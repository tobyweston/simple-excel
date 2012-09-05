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

package bad.robot.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;
import java.text.ParseException;

import static bad.robot.excel.DateUtil.createDate;
import static bad.robot.excel.RowBuilder.aRow;
import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.matchers.WorkbookMatcher.sameWorkbook;
import static bad.robot.excel.valuetypes.ColumnIndex.column;
import static bad.robot.excel.valuetypes.Coordinate.coordinate;
import static bad.robot.excel.valuetypes.ExcelColumnIndex.*;
import static bad.robot.excel.valuetypes.RowIndex.row;
import static java.util.Calendar.FEBRUARY;
import static java.util.Calendar.MAY;
import static org.hamcrest.MatcherAssert.assertThat;

public class PoiWorkbookMutatorTest {

    @Test(expected = IllegalArgumentException.class)
    public void shouldNotAllowNullWorkbook() {
        new PoiWorkbookMutator(null);
    }

    @Test
    public void shouldReplaceCell() throws IOException {
        Workbook workbook = getWorkbook("shouldReplaceCellTemplate.xls");
        new PoiWorkbookMutator(workbook).replaceCell(coordinate(column(A), row(1)), "Hello World");

        assertThat(workbook, sameWorkbook(getWorkbook("shouldReplaceCellTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceDateCell() throws IOException {
        Workbook workbook = getWorkbook("shouldReplaceDateCellTemplate.xls");
        new PoiWorkbookMutator(workbook).replaceCell(coordinate(column(A), row(1)), createDate(22, MAY, 1997));

        assertThat(workbook, sameWorkbook(getWorkbook("shouldReplaceDateCellTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceCellsInComplicatedExample() throws IOException {
        Workbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
        new PoiWorkbookMutator(workbook)
                .replaceCell(coordinate(column(C), row(5)), "Very")
                .replaceCell(coordinate(column(D), row(11)), "Complicated")
                .replaceCell(coordinate(column(G), row(3)), "Example")
                .replaceCell(coordinate(column(H), row(7)), "Of")
                .replaceCell(coordinate(column(J), row(10)), "Templated")
                .replaceCell(coordinate(column(M), row(15)), "Spreadsheet");

        assertThat(workbook, sameWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls")));
    }

    @Test
    public void shouldReplaceCellsInComplicatedAlternateSyntaxExample() throws IOException {
        Workbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
        new PoiWorkbookMutator(workbook)
                .replaceCell(coordinate(C, 5), "Very")
                .replaceCell(coordinate(D, 11), "Complicated")
                .replaceCell(coordinate(G, 3), "Example")
                .replaceCell(coordinate(H, 7), "Of")
                .replaceCell(coordinate(J, 10), "Templated")
                .replaceCell(coordinate(M, 15), "Spreadsheet");

        assertThat(workbook, sameWorkbook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls")));
    }

    @Test
    public void shouldAppendRow() throws IOException, ParseException {
        Workbook workbook = getWorkbook("shouldAppendRowTemplate.xls");
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
        new PoiWorkbookMutator(workbook).appendRowToFirstSheet(row.build());

        assertThat(workbook, sameWorkbook(getWorkbook("shouldAppendRowTemplateExpected.xls")));
    }

}