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

import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.Date;

import static bad.robot.excel.RowBuilder.aRow;
import static bad.robot.excel.matchers.WorkbookEqualityMatcher.sameWorkBook;
import static bad.robot.excel.valuetypes.CellIndex.cellIndex;
import static bad.robot.excel.valuetypes.ColumnIndex.column;
import static bad.robot.excel.valuetypes.Coordinate.coordinate;
import static bad.robot.excel.valuetypes.ExcelColumnIndex.*;
import static bad.robot.excel.valuetypes.RowIndex.row;
import static org.apache.commons.lang3.time.DateUtils.parseDateStrictly;
import static org.hamcrest.Matchers.is;
import static org.junit.Assert.assertThat;

public class PoiWorkbookMutatorTest {

    @Test(expected = IllegalArgumentException.class)
    public void shouldNotAllowNullWorkbook() {
        new PoiWorkbookMutator(null);
    }

    @Test
    public void shouldReplaceCell() throws IOException {
        HSSFWorkbook workbook = getWorkbook("shouldReplaceCellTemplate.xls");
        new PoiWorkbookMutator(workbook).replaceCell(coordinate(column(A), row(1)), "Hello World");

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldReplaceCellTemplateExpected.xls"))));
    }

    @Test
    public void shouldReplaceCellsInComplicatedExample() throws IOException {
        HSSFWorkbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
        new PoiWorkbookMutator(workbook)
                .replaceCell(coordinate(column(C), row(5)), "Very")
                .replaceCell(coordinate(column(D), row(11)), "Complicated")
                .replaceCell(coordinate(column(G), row(3)), "Example")
                .replaceCell(coordinate(column(H), row(7)), "Of")
                .replaceCell(coordinate(column(J), row(10)), "Templated")
                .replaceCell(coordinate(column(M), row(15)), "Spreadsheet");

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls"))));
    }

    @Test
    public void shouldReplaceCellsInComplicatedAlternateSyntaxExample() throws IOException {
        HSSFWorkbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
        new PoiWorkbookMutator(workbook)
                .replaceCell(coordinate(C, 5), "Very")
                .replaceCell(coordinate(D, 11), "Complicated")
                .replaceCell(coordinate(G, 3), "Example")
                .replaceCell(coordinate(H, 7), "Of")
                .replaceCell(coordinate(J, 10), "Templated")
                .replaceCell(coordinate(M, 15), "Spreadsheet");

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls"))));
    }

    @Test
    public void shouldAppendRow() throws IOException, ParseException {
        HSSFWorkbook workbook = getWorkbook("shouldAppendRowTemplate.xls");
        RowBuilder row = aRow()
                .withString(cellIndex(A), "This")
                .withString(cellIndex(C), "Row")
                .withString(cellIndex(D), "was")
                .withString(cellIndex(E), "inserted")
                .withString(cellIndex(H), "here")
                .withString(cellIndex(I), "by")
                .withString(cellIndex(J), "some")
                .withString(cellIndex(K), "smart")
                .withString(cellIndex(N), "Programmer")
                .withInteger(cellIndex(L), 1)
                .withDate(cellIndex(P), fromSimpleString("79-02-11"))
                .withDouble(cellIndex(O), new Double("0.123456789"));
        new PoiWorkbookMutator(workbook).appendRowToFirstSheet(row.build());

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldAppendRowTemplateExpected.xls"))));
    }

    private static HSSFWorkbook getWorkbook(String file) throws IOException {
        InputStream stream = PoiWorkbookMutatorTest.class.getResourceAsStream(file);
        if (stream == null)
            throw new FileNotFoundException(file);
        return new HSSFWorkbook(stream);
    }

    public static Date fromSimpleString(String date) throws ParseException {
        return parseDateStrictly(date, new String[]{FastDateFormat.getInstance("yy-MM-dd").getPattern()});
    }

}