/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
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

import bad.robot.excel.cell.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.WorkbookResource.getCellForCoordinate;
import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.column.ExcelColumnIndex.B;
import static bad.robot.excel.column.ExcelColumnIndex.C;
import static bad.robot.excel.matchers.CellType.adaptPoi;
import static bad.robot.excel.sheet.Coordinate.coordinate;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.instanceOf;
import static org.hamcrest.Matchers.is;

public class CellTypeTest {

    private Workbook workbook;

    @Before
    public void loadWorkbook() throws IOException {
        workbook = getWorkbook("cellTypes.xls");
    }

    @Test
    public void adaptBooleanCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 6), workbook)), is(instanceOf(BooleanCell.class)));
    }

    @Test
    public void adaptFormulaCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 7), workbook)), is(instanceOf(FormulaCell.class)));
    }

    @Test
    public void adaptFormulaTypeToError() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 8), workbook)), is(instanceOf(ErrorCell.class)));
    }

    @Test
    public void adaptNumericTypeToDoubleCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 2), workbook)), is(instanceOf(DoubleCell.class)));
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 3), workbook)), is(instanceOf(DoubleCell.class)));
    }

    @Test
    public void adaptNumericTypeToDateCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 1), workbook)), is(instanceOf(DateCell.class)));
    }

    @Test
    public void adaptStringCell() throws Exception {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 4), workbook)), is(instanceOf(StringCell.class)));
    }

    @Test
    public void adaptHyperlinkCell() throws Exception {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 5), workbook)), is(instanceOf(HyperlinkCell.class)));
    }

    @Test
    public void adaptInternalLinkCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(B, 9), workbook)), is(instanceOf(StringCell.class)));
    }

    @Test
    public void adaptBlankCell() throws IOException {
        assertThat(adaptPoi(getCellForCoordinate(coordinate(C, 1), workbook)), is(instanceOf(BlankCell.class)));
    }
}
