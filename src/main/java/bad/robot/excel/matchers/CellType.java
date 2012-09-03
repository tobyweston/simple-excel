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

import static org.apache.poi.ss.usermodel.Cell.*;

public enum CellType implements CellAdapter {

    Boolean(CELL_TYPE_BOOLEAN) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new BooleanCell(cell.getBooleanCellValue());
        }
    },
    Error(CELL_TYPE_ERROR) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new ErrorCell(cell.getErrorCellValue());
        }
    },
    Formula(CELL_TYPE_FORMULA) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            if (cell.getCachedFormulaResultType() == CELL_TYPE_ERROR)
                return new ErrorCell(cell.getErrorCellValue());
            return new FormulaCell(cell.getCellFormula());
        }
    },
    Numeric(CELL_TYPE_NUMERIC) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new DoubleCell(cell.getNumericCellValue());
        }
    },
    String(CELL_TYPE_STRING) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new StringCell(cell.getStringCellValue());
        }
    },
    Blank(CELL_TYPE_BLANK) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new BlankCell();
        }
    };

    private final Integer poiType;

    CellType(Integer poiType) {
        this.poiType = poiType;
    }

    private static CellAdapter getAdapterFor(org.apache.poi.ss.usermodel.Cell poi) {
        if (poi == null)
            return Blank;

        for (CellType type : values()) {
            if (type.poiType == poi.getCellType())
                return type;
        }
        throw new RuntimeException("Unknown poi type " + poi.getCellType());
    }

    public static Cell adaptPoi(org.apache.poi.ss.usermodel.Cell cell) {
        return getAdapterFor(cell).adapt(cell);
    }

}
