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

package bad.robot.excel.column;

import bad.robot.excel.AbstractValueType;

public class ColumnIndex extends AbstractValueType<Integer> {

    public static ColumnIndex column(ExcelColumnIndex index) {
        return new ColumnIndex(index.ordinal());
    }

    public static ColumnIndex column(String index) {
        return column(ExcelColumnIndex.getColumn(index));
    }

    private ColumnIndex(Integer value) {
        super(value);
    }
}
