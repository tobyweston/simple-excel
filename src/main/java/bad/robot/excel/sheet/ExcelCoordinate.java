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

package bad.robot.excel.sheet;

import bad.robot.excel.column.ExcelColumnIndex;
import bad.robot.excel.row.ExcelRowIndex;

public class ExcelCoordinate {

    private final ExcelColumnIndex column;
    private final ExcelRowIndex row;

    public static ExcelCoordinate coordinate(ExcelColumnIndex column, int row) {
        return new ExcelCoordinate(column, ExcelRowIndex.row(row));
    }

    public static ExcelCoordinate coordinate(ExcelColumnIndex column, ExcelRowIndex row) {
        return new ExcelCoordinate(column, row);
    }

    public static ExcelCoordinate coordinate(String column, int row) {
        return coordinate(ExcelColumnIndex.getColumn(column), row);
    }

    public static ExcelCoordinate coordinate(String column, ExcelRowIndex row) {
    	return coordinate(ExcelColumnIndex.getColumn(column), row);
    }

    public ExcelCoordinate(ExcelColumnIndex column, ExcelRowIndex row) {
        this.column = column;
        this.row = row;
    }

}
