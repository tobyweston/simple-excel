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

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.hamcrest.Description;
import org.hamcrest.Matcher;
import org.hamcrest.TypeSafeMatcher;

import static bad.robot.excel.PoiToExcelCoordinateCoercions.asExcelCoordinate;
import static bad.robot.excel.matchers.SheetNumberMatcher.hasSameNumberOfSheetsAs;
import static java.lang.String.format;

public class WorkbookEqualityMatcher extends TypeSafeMatcher<Workbook> {

    private final Workbook expectedWorkbook;
    private String lastError;

    WorkbookEqualityMatcher(Workbook expectedWorkbook) {
        this.expectedWorkbook = expectedWorkbook;
    }

    public static Matcher<Workbook> sameWorkBook(Workbook expectedWorkbook) {
        return new WorkbookEqualityMatcher(expectedWorkbook);
    }

    @Override
    public boolean matchesSafely(Workbook actual) {
        try {
            if (!hasSameNumberOfSheetsAs(expectedWorkbook).matches(actual))
                return false;

            for (int a = 0; a < actual.getNumberOfSheets(); a++) {
                Sheet actualSheet = actual.getSheetAt(a);
                Sheet expectedSheet = expectedWorkbook.getSheetAt(a);

                if (!RowNumberMatcher.hasSameNumberOfRowAs(expectedSheet).matches(actualSheet))
                    return false;

                for (int i = 0; i <= expectedSheet.getLastRowNum(); i++)
                    checkIfRowEqual(actualSheet, expectedSheet, i);
            }
        } catch (WorkbookDiscrepancyException e) {
            lastError = e.getMessage();
            return false;
        }
        return true;
    }

    private void checkIfRowEqual(Sheet actualSheet, Sheet expectedSheet, int i) throws WorkbookDiscrepancyException {
        org.apache.poi.ss.usermodel.Row expectedRow = expectedSheet.getRow(i);
        org.apache.poi.ss.usermodel.Row actualRow = actualSheet.getRow(i);
        if (bothRowsAreNull(expectedRow, actualRow))
            return;
        if (oneRowIsNullAndOtherNot(expectedRow, actualRow))
            throw new WorkbookDiscrepancyException("One of rows was null");
        if (expectedRow.getLastCellNum() != actualRow.getLastCellNum())
            throw new WorkbookDiscrepancyException(format("Different number of cells: expected: '%d' actual '%d'", expectedRow.getLastCellNum(), actualRow.getLastCellNum()));

        for (int j = 0; j <= expectedRow.getLastCellNum(); j++)
            checkIfCellEqual(expectedRow, actualRow, j);
    }

    private void checkIfCellEqual(org.apache.poi.ss.usermodel.Row expectedRow, org.apache.poi.ss.usermodel.Row actualRow, int j) throws WorkbookDiscrepancyException {
        org.apache.poi.ss.usermodel.Cell expectedCell = expectedRow.getCell(j);
        org.apache.poi.ss.usermodel.Cell actualCell = actualRow.getCell(j);
        if (bothCellsAreNull(expectedCell, actualCell))
            return;

        if (bothCellsAreNullOrBlank(expectedCell, actualCell))
            return;

        if (anyOfTheCellsAreNull(expectedCell, actualCell))
            throw new WorkbookDiscrepancyException("One of cells was null");

        CellType expectedCellType = CellType.valueOf(expectedCell.getCellType());
        CellType actualCellType = CellType.valueOf(actualCell.getCellType());

        if (expectedCellType != actualCellType)
            throw new WorkbookDiscrepancyException(format("Cell at %s has different types: expected: '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCellType, actualCellType));

        expectedCellType.assertSameValue(expectedCell, actualCell);

    }

    private boolean cellIsNullOrBlank(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null || cell.getCellType() == 3)
            return true;
        return false;
    }

    private boolean bothCellsAreNullOrBlank(org.apache.poi.ss.usermodel.Cell expected, org.apache.poi.ss.usermodel.Cell actual) {
        return cellIsNullOrBlank(expected) && cellIsNullOrBlank(actual);
    }

    private boolean oneRowIsNullAndOtherNot(org.apache.poi.ss.usermodel.Row expectedRow, org.apache.poi.ss.usermodel.Row actualRow) {
        if (actualRow == null || expectedRow == null)
            return true;
        return false;
    }

    private boolean bothRowsAreNull(org.apache.poi.ss.usermodel.Row expectedRow, org.apache.poi.ss.usermodel.Row actualRow) {
        if ((actualRow == null && expectedRow == null))
            return true;
        return false;
    }

    private boolean anyOfTheCellsAreNull(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) {
        return actualCell == null || expectedCell == null;
    }

    private boolean bothCellsAreNull(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) {
        return (actualCell == null && expectedCell == null);
    }

    @Override
    public void describeTo(Description description) {
        description.appendText(lastError);
    }


}
