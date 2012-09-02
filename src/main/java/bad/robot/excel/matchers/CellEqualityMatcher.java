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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.PoiToExcelCoercions.asExcelCoordinate;
import static bad.robot.excel.PoiToExcelCoercions.asExcelRow;
import static bad.robot.excel.matchers.CellType.adaptPoi;

public class CellEqualityMatcher extends TypeSafeDiagnosingMatcher<Row> {

    private final Row expected;

    private CellEqualityMatcher(Row expected) {
        this.expected = expected;
    }

    public static CellEqualityMatcher hasSameCellsAs(Row expected) {
        return new CellEqualityMatcher(expected);
    }

    @Override
    protected boolean matchesSafely(Row actual, Description mismatch) {
        for (Cell cell : expected)
            if (!equal(cell, actual, mismatch))
                return false;
        return true;
    }

    private boolean equal(Cell expectedPoi, Row actualPoi, Description mismatch) {
        bad.robot.excel.Cell expected = adaptPoi(expectedPoi);
        bad.robot.excel.Cell actual = adaptPoi(actualPoi.getCell(expectedPoi.getColumnIndex()));

        if (!expected.equals(actual)) {
            mismatch.appendText("cell at ").appendValue(asExcelCoordinate(expectedPoi)).appendText(" contained ").appendValue(actual).appendText(" expected ").appendValue(expected);
            return false;
        }
        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendText("equality of all cells on row ").appendValue(asExcelRow(expected));
    }

}
