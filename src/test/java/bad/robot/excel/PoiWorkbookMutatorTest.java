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
        new PoiWorkbookMutator(workbook).replaceCell(coordinate(column(0), row(0)), "Hello World");

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldReplaceCellTemplateExpected.xls"))));
    }

    @Test
    public void shouldReplaceCellsInComplicatedExample() throws IOException {
        HSSFWorkbook workbook = getWorkbook("shouldReplaceCellsInComplicatedExampleTemplate.xls");
        new PoiWorkbookMutator(workbook)
                .replaceCell(coordinate(column(2), row(4)), "Very")
                .replaceCell(coordinate(column(3), row(10)), "Complicated")
                .replaceCell(coordinate(column(6), row(2)), "Example")
                .replaceCell(coordinate(column(7), row(6)), "Of")
                .replaceCell(coordinate(column(9), row(9)), "Templated")
                .replaceCell(coordinate(column(12), row(14)), "Spreadsheet");

        assertThat(workbook, is(sameWorkBook(getWorkbook("shouldReplaceCellsInComplicatedExampleTemplateExpected.xls"))));
    }

    @Test
    public void shouldAppendRow() throws IOException, ParseException {
        HSSFWorkbook workbook = getWorkbook("shouldAppendRowTemplate.xls");
        RowBuilder row = aRow()
                .withString(cellIndex(0), "This")
                .withString(cellIndex(2), "Row")
                .withString(cellIndex(3), "was")
                .withString(cellIndex(4), "inserted")
                .withString(cellIndex(7), "here")
                .withString(cellIndex(8), "by")
                .withString(cellIndex(9), "some")
                .withString(cellIndex(10), "smart")
                .withString(cellIndex(13), "Programmer")
                .withInteger(cellIndex(11), 1)
                .withDate(cellIndex(15), fromSimpleString("79-02-11"))
                .withDouble(cellIndex(14), new Double("0.123456789"));
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