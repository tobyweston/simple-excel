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

import bad.robot.excel.*;
import bad.robot.excel.valuetypes.Coordinate;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.PoiToExcelCoercions.asExcelRow;
import static bad.robot.excel.matchers.CellType.adaptPoi;
import static bad.robot.excel.valuetypes.Coordinate.coordinate;
import static bad.robot.excel.valuetypes.ExcelColumnIndex.B;
import static bad.robot.excel.valuetypes.ExcelColumnIndex.C;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.instanceOf;
import static org.hamcrest.Matchers.is;

public class CellTypeTest {

    @Test
    public void adaptBooleanCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 6))), is(instanceOf(BooleanCell.class)));
    }

    @Test
    public void adaptFormulaCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 7))), is(instanceOf(FormulaCell.class)));
    }

    @Test
    public void adaptFormulaTypeToError() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 8))), is(instanceOf(ErrorCell.class)));
    }

    @Test
    public void adaptNumericTypeToDoubleCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 2))), is(instanceOf(DoubleCell.class)));
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 3))), is(instanceOf(DoubleCell.class)));
    }

    @Test
    public void adaptNumericTypeToDateCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 1))), is(instanceOf(DateCell.class)));
    }

    @Test
    public void adaptStringCell() throws Exception {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 4))), is(instanceOf(StringCell.class)));
    }

    @Test
    public void adaptHyperlinkCell() throws Exception {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 5))), is(instanceOf(HyperlinkCell.class)));
    }

    @Test
    public void adaptBlankCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(C, 1))), is(instanceOf(BlankCell.class)));
    }

    private static org.apache.poi.ss.usermodel.Cell getCellForCoordinate(Coordinate coordinate) throws IOException {
        org.apache.poi.ss.usermodel.Row row = getRowForCoordinate(coordinate);
        return row.getCell(coordinate.getColumn().value());
    }

    private static org.apache.poi.ss.usermodel.Row getRowForCoordinate(Coordinate coordinate) throws IOException {
        Workbook workbook = WorkbookResource.getWorkbook("cellTypes.xls");
        Sheet sheet = workbook.getSheetAt(coordinate.getSheet().value());
        org.apache.poi.ss.usermodel.Row row = sheet.getRow(coordinate.getRow().value());
        if (row == null)
            throw new IllegalStateException("expected to find a row at " + asExcelRow(row));
        return row;
    }

}
