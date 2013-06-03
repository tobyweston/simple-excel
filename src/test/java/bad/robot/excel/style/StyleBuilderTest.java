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

package bad.robot.excel.style;

import bad.robot.excel.cell.Cell;
import bad.robot.excel.cell.DataFormat;
import bad.robot.excel.cell.DoubleCell;
import bad.robot.excel.column.ColumnIndex;
import bad.robot.excel.row.Row;
import bad.robot.excel.workbook.PoiWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.util.HashMap;

import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.cell.DataFormat.asTwoDecimalPlacesNumber;
import static bad.robot.excel.column.ColumnIndex.column;
import static bad.robot.excel.column.ExcelColumnIndex.A;
import static bad.robot.excel.matchers.WorkbookMatcher.sameWorkbook;
import static bad.robot.excel.style.Border.border;
import static bad.robot.excel.style.BorderStyle.None;
import static bad.robot.excel.style.BorderStyle.ThinSolid;
import static bad.robot.excel.style.BottomBorder.bottom;
import static bad.robot.excel.style.LeftBorder.left;
import static bad.robot.excel.style.RightBorder.right;
import static bad.robot.excel.style.StyleBuilder.aStyle;
import static bad.robot.excel.style.TopBorder.top;
import static org.hamcrest.MatcherAssert.assertThat;

public class StyleBuilderTest {

    private final Border border = border(top(None), right(None), bottom(ThinSolid), left(ThinSolid));
    private final DataFormat numberFormat = asTwoDecimalPlacesNumber();

    @Test
    public void exampleOfCreatingARow() throws Exception {
        Cell cell = new DoubleCell(9999.99d, aStyle().with(border).with(numberFormat));

        HashMap<ColumnIndex, Cell> columns = new HashMap<ColumnIndex, Cell>();
        columns.put(column(A), cell);

        Row row = new Row(columns);

        Workbook workbook = getWorkbook("emptySheet.xls");
        new PoiWorkbook(workbook).appendRowToFirstSheet(row);

        assertThat(workbook, sameWorkbook(getWorkbook("sheetWithSingleCell.xls")));
    }

}
