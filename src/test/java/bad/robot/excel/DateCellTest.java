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

package bad.robot.excel;

import bad.robot.excel.valuetypes.Coordinate;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;

import static bad.robot.excel.DateUtil.createDate;
import static bad.robot.excel.WorkbookResource.*;
import static bad.robot.excel.matchers.CellMatcher.equalTo;
import static bad.robot.excel.valuetypes.Coordinate.coordinate;
import static bad.robot.excel.valuetypes.ExcelColumnIndex.B;
import static java.util.Calendar.MARCH;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.is;
import static org.junit.Assert.fail;

public class DateCellTest {

    @Test
    @Ignore
    public void shouldSetDataFormatWhenAddingACell() {
        fail();
    }

    @Test
    @Ignore
    public void shouldSetDataFormatWhenReplacingAnUnformattedCell() {
        fail();
    }

    @Test
    @Ignore
    public void shouldNotSetDataFormatWhenReplacingAFormattedCell() {
        fail();
    }

    @Test
    public void replaceNonDateCellWithACell() throws IOException {
        Workbook workbook = getWorkbook("cellTypes.xls");
        PoiWorkbookMutator sheet = new PoiWorkbookMutator(workbook);
        Coordinate coordinate = coordinate(B, 2);

        assertThat(getCellForCoordinate(coordinate, workbook), equalTo(new DoubleCell(1001d)));
        sheet.replaceCell(coordinate, createDate(15, MARCH, 2012));

        assertThat(getCellForCoordinate(coordinate, workbook), equalTo(new DateCell(createDate(15, MARCH, 2012))));
        assertThat(getCellDataFormatAtCoordinate(coordinate, workbook), is("dd-MMM-yyyy"));
        assertThat("should not have affected a shared data format", getCellDataFormatAtCoordinate(coordinate(B, 7), workbook), is("General"));
    }
}
