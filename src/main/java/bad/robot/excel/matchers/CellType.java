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

import static bad.robot.excel.PoiToExcelCoordinateCoercions.asExcelCoordinate;
import static java.lang.String.format;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.apache.poi.ss.usermodel.Cell.*;

enum CellType {
    Boolean(CELL_TYPE_BOOLEAN) {
        @Override
        public void assertSameValue(Cell expectedCell, Cell actualCell) throws WorkbookDiscrepancyException {
            if (expectedCell.getBooleanCellValue() != actualCell.getBooleanCellValue())
                throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getBooleanCellValue(), actualCell.getBooleanCellValue()));

        }
    },
    Error(CELL_TYPE_ERROR) {
        @Override
        public void assertSameValue(Cell expectedCell, Cell actualCell) throws WorkbookDiscrepancyException {
            if (expectedCell.getErrorCellValue() != actualCell.getErrorCellValue()) {
                throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getErrorCellValue(), actualCell.getErrorCellValue()));
            }

        }
    },
    Formula(CELL_TYPE_FORMULA) {
        @Override
        public void assertSameValue(Cell expectedCell, Cell actualCell) throws WorkbookDiscrepancyException {
            if (!expectedCell.getCellFormula().equals(actualCell.getCellFormula())) {
                throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getCellFormula(), actualCell.getCellFormula()));
            }

        }
    },
    Numeric(CELL_TYPE_NUMERIC) {
        @Override
        public void assertSameValue(Cell expectedCell, Cell actualCell) throws WorkbookDiscrepancyException {
            if (expectedCell.getNumericCellValue() != actualCell.getNumericCellValue())
                throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expectedCell), expectedCell.getNumericCellValue(), actualCell.getNumericCellValue()));
        }
    },
    String(CELL_TYPE_STRING, CELL_TYPE_BLANK) {
        @Override
        public void assertSameValue(Cell expected, Cell actual) throws WorkbookDiscrepancyException {
            if (actual.getCellType() == CELL_TYPE_BLANK && isBlank(expected.getStringCellValue()))
                return;
            if (!expected.getStringCellValue().equals(actual.getStringCellValue()))
                throw new WorkbookDiscrepancyException(format("Cell at %s has different values: expected '%s' actual '%s'", asExcelCoordinate(expected), expected.getStringCellValue(), actual.getStringCellValue()));
        }
    };

    private Integer[] poiValues;

    CellType(Integer... poiValues) {
        this.poiValues = poiValues;
    }

    static CellType valueOf(int cellType) {
        for (CellType type : values()) {
            for (Integer poiValue : type.poiValues) {
                if (poiValue == cellType)
                    return type;
            }
        }
        throw new RuntimeException("Unknown poi type " + cellType);
    }

    public abstract void assertSameValue(Cell expectedCell, Cell actualCell) throws WorkbookDiscrepancyException;
}
