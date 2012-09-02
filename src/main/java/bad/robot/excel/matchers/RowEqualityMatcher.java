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
import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.PoiToExcelCoercions.asExcelCoordinate;
import static bad.robot.excel.matchers.CellNumberMatcher.hasSameNumberOfCellsAs;
import static java.lang.String.format;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK;

public class RowEqualityMatcher extends TypeSafeDiagnosingMatcher<Sheet> {

    private final Sheet expected;

    public static RowEqualityMatcher rowsEqual(Sheet expected) {
        return new RowEqualityMatcher(expected);
    }

    private RowEqualityMatcher(Sheet expected) {
        this.expected = expected;
    }

    @Override
    protected boolean matchesSafely(Sheet actual, Description mismatch) {
        try {
            for (Row row : expected)
                verify(row, actual.getRow(row.getRowNum()), mismatch);
        } catch (WorkbookDiscrepancyException e) {
            mismatch.appendText(e.getMessage());
            return false;
        }
        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendText("equality on all rows in ").appendValue(expected.getSheetName());
    }

    private void verify(Row expected, Row actual, Description mismatch) throws WorkbookDiscrepancyException {
        if (isBothNull(expected, actual)) // is this needed?
            return;

        if (expectedRowIsMissingFrom(actual))
            throw new WorkbookDiscrepancyException(format("row %s is missing", expected.getRowNum() + 1));

        if (!hasSameNumberOfCellsAs(expected).matches(actual))
            throw new WorkbookDiscrepancyException(format("Different number of cells: expected: '%d' actual '%d'", expected.getLastCellNum(), actual.getLastCellNum()));

        for (Cell cell : expected)
            verify(cell, actual.getCell(cell.getColumnIndex()), mismatch);
    }

    private void verify(Cell expected, Cell actual, Description mismatch) throws WorkbookDiscrepancyException {
        bad.robot.excel.Cell expectedCell = CellType.adaptPoi(expected);
        bad.robot.excel.Cell actualCell = CellType.adaptPoi(actual);

        if (!expectedCell.equals(actualCell)) {
            mismatch.appendText("cell at ").appendValue(asExcelCoordinate(expected)).appendText(" contained ").appendValue(actualCell).appendText(" expected ").appendValue(expectedCell);
            throw new WorkbookDiscrepancyException("");
        }

        // deprecated below this line

//        if (isBothNull(expected, actual))
//            return;
//
//        if (bothCellsAreNullOrBlank(expected, actual))
//            return;
//
//        if (anyOfTheCellsAreNull(expected, actual))
//            throw new WorkbookDiscrepancyException("One of cells was null");
//
//        DeprecatedCellType expectedCellType = DeprecatedCellType.valueOf(expected.getCellType());
//        DeprecatedCellType actualCellType = DeprecatedCellType.valueOf(actual.getCellType());
//
//        if (expectedCellType != actualCellType) // DONE with equality
//            throw new WorkbookDiscrepancyException(format("Cell at %s has different types: expected: '%s' actual '%s'", asExcelCoordinate(expected), expectedCellType, actualCellType));
//
//        expectedCellType.assertSameValue(expected, actual);
    }

    private boolean isBothNull(Object first, Object second) {
        return first == null && second == null;
    }

    private boolean expectedRowIsMissingFrom(Row second) {
        return second == null;
    }

    private boolean bothCellsAreNullOrBlank(Cell expected, Cell actual) {
        return cellIsNullOrBlank(expected) && cellIsNullOrBlank(actual);
    }

    private boolean anyOfTheCellsAreNull(Cell expectedCell, Cell actualCell) {
        return actualCell == null || expectedCell == null;
    }

    private boolean cellIsNullOrBlank(Cell cell) {
        return cell == null || cell.getCellType() == CELL_TYPE_BLANK;
    }

}
