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

import bad.robot.excel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.PoiToExcelCoercions.asExcelCoordinate;
import static bad.robot.excel.matchers.CellType.adaptPoi;

public class CellMatcher extends TypeSafeDiagnosingMatcher<Row> {

    private final Cell expected;
    private final int columnIndex;
    private final String coordinate;

    private CellMatcher(org.apache.poi.ss.usermodel.Cell expectedPoi) {
        this.expected = adaptPoi(expectedPoi);
        this.coordinate = asExcelCoordinate(expectedPoi);
        this.columnIndex = expectedPoi.getColumnIndex();
    }

    public static CellMatcher hasSameCellAs(org.apache.poi.ss.usermodel.Cell expected) {
        return new CellMatcher(expected);
    }

    @Override
    protected boolean matchesSafely(Row row, Description mismatch) {
        Cell actual = adaptPoi(row.getCell(columnIndex));

        if (!expected.equals(actual)) {
            mismatch.appendText("cell at ").appendValue(coordinate).appendText(" contained ").appendValue(actual).appendText(" expected ").appendValue(expected);
            return false;
        }
        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendText("equality of cell ").appendValue(coordinate);
    }
}
