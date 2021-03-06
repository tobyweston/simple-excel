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

import bad.robot.excel.AbstractValueType;

public class SheetIndex extends AbstractValueType<Integer> {

    public static SheetIndex sheet(Integer value) {
        if (value <= 0)
            throw new IllegalArgumentException("sheet indices start at 1");
        return new SheetIndex(value - 1);
    }

    private SheetIndex(Integer value) {
        super(value);
    }
}
