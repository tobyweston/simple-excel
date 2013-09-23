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

package bad.robot.excel.row;

import bad.robot.excel.column.ColumnIndex;
import bad.robot.excel.style.Style;

import java.util.Date;

public interface StyledRowBuilder {
    StyledRowBuilder withBlank(ColumnIndex index, Style style);

    StyledRowBuilder withString(ColumnIndex index, String text, Style style);

    StyledRowBuilder withDouble(ColumnIndex index, Double value, Style style);

    StyledRowBuilder withInteger(ColumnIndex index, Integer value, Style style);

    StyledRowBuilder withDate(ColumnIndex index, Date date, Style style);

    StyledRowBuilder withFormula(ColumnIndex index, String formula, Style style);
}
