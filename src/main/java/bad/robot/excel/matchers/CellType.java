/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package bad.robot.excel.matchers;


import bad.robot.excel.cell.*;

import static bad.robot.excel.PoiToExcelCoercions.asExcelCoordinate;
import static bad.robot.excel.cell.Hyperlink.hyperlink;
import static org.apache.poi.ss.usermodel.CellType.*;
import static org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted;

public enum CellType implements CellAdapter {

    Boolean(BOOLEAN) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new BooleanCell(cell.getBooleanCellValue());
        }
    },
    Error(ERROR) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new ErrorCell(cell.getErrorCellValue());
        }
    },
    Formula(FORMULA) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            if (cell.getCachedFormulaResultType() == ERROR)
                return new ErrorCell(cell.getErrorCellValue());
            return new FormulaCell(cell.getCellFormula());
        }
    },
    Numeric(NUMERIC) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            if (isCellDateFormatted(cell))
                return new DateCell(cell.getDateCellValue());
            return new DoubleCell(cell.getNumericCellValue());
        }
    },
    String(STRING) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            if (cell.getHyperlink() != null && containsUrl(cell.getHyperlink()))
                return new HyperlinkCell(hyperlink(cell.getStringCellValue(), cell.getHyperlink().getAddress()));
            
            if (cell.getStringCellValue() == null || "".equals(cell.getStringCellValue()))
                return new BlankCell();
            
            return new StringCell(cell.getStringCellValue());
        }

        private boolean containsUrl(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
            return hyperlink.getAddress().startsWith("http://") || hyperlink.getAddress().startsWith("file://");
        }
    },
    Blank(BLANK) {
        @Override
        public Cell adapt(org.apache.poi.ss.usermodel.Cell cell) {
            return new BlankCell();
        }
    };

    private final org.apache.poi.ss.usermodel.CellType poiType;

    CellType(org.apache.poi.ss.usermodel.CellType poiType) {
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
        try {
            return getAdapterFor(cell).adapt(cell);
        } catch (Exception e) {
            throw new IllegalStateException("problem with cell " + asExcelCoordinate(cell), e);
        }
    }

}
