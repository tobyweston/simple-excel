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
import static java.lang.String.format;
import static org.apache.commons.lang3.StringUtils.isBlank;

public class WorkbookEqualityMatcher extends TypeSafeMatcher<Workbook> {

    private final Workbook expectedWorkbook;
    private String lastError;

    WorkbookEqualityMatcher(Workbook expectedWorkbook) {
        this.expectedWorkbook = expectedWorkbook;
    }

    public static Matcher<Workbook> sameWorkBook(Workbook expectedWorkbook) {
        return new WorkbookEqualityMatcher(expectedWorkbook);
    }

    enum CellType {
        BOOLEAN(4) {
            @Override
            public void assertSameValue(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) throws WorkbookDiscrepancyException {
                if (expectedCell.getBooleanCellValue() != actualCell.getBooleanCellValue())
                    throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getBooleanCellValue(), actualCell.getBooleanCellValue()));

            }
        },
        ERROR(5) {
            @Override
            public void assertSameValue(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) throws WorkbookDiscrepancyException {
                if (expectedCell.getErrorCellValue() != actualCell.getErrorCellValue()) {
                    throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getErrorCellValue(), actualCell.getErrorCellValue()));
                }

            }
        },
        FORMULA(2) {
            @Override
            public void assertSameValue(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) throws WorkbookDiscrepancyException {
                if (!expectedCell.getCellFormula().equals(actualCell.getCellFormula())) {
                    throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getCellFormula(), actualCell.getCellFormula()));
                }

            }
        },
        NUMERIC(0) {
            @Override
            public void assertSameValue(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) throws WorkbookDiscrepancyException {
                if (expectedCell.getNumericCellValue() != actualCell.getNumericCellValue())
                    throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getNumericCellValue(), actualCell.getNumericCellValue()));
            }
        },
        STRING(1, 3) {
            @Override
            public void assertSameValue(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) throws WorkbookDiscrepancyException {
                if (actualCell.getCellType() == 3 && isBlank(expectedCell.getStringCellValue())) {
                    return;
                }
                if (!expectedCell.getStringCellValue().equals(actualCell.getStringCellValue())) {
                    throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getStringCellValue(), actualCell.getStringCellValue()));
                }
            }
        };

        private Integer[] poiValues;

        CellType(Integer... poiValues) {
            this.poiValues = poiValues;
        }

        public static CellType valueOf(int cellType) {
            for (CellType type : values()) {
                for (Integer poiValue : type.poiValues) {
                    if (poiValue == cellType)
                        return type;
                }
            }
            throw new RuntimeException("Unknown poi type " + cellType);
        }

        public abstract void assertSameValue(org.apache.poi.ss.usermodel.Cell expectedCell, org.apache.poi.ss.usermodel.Cell actualCell) throws WorkbookDiscrepancyException;
    }

    @Override
    public boolean matchesSafely(Workbook actual) {
        try {
            if (actual.getNumberOfSheets() != expectedWorkbook.getNumberOfSheets())
                throw new WorkbookDiscrepancyException(format("Different number of sheets: expected: '%d' actual '%d'", expectedWorkbook.getNumberOfSheets(), ((Workbook) actual).getNumberOfSheets()));

            for (int a = 0; a < actual.getNumberOfSheets(); a++) {
                Sheet actualSheet = actual.getSheetAt(a);
                Sheet expectedSheet = expectedWorkbook.getSheetAt(a);

                if (actualSheet.getLastRowNum() != expectedSheet.getLastRowNum())
                    throw new WorkbookDiscrepancyException(format("Different number of rows: expected: '%d' actual '%d'", expectedSheet.getLastRowNum(), actualSheet.getLastRowNum()));

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


    static class WorkbookDiscrepancyException extends Exception {
        public WorkbookDiscrepancyException(String message) {
            super(message);
        }
    }
}
