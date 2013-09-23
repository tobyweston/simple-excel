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

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Iterator;

public class SheetIterable implements Iterable<Sheet> {

    private final Workbook workbook;

    private SheetIterable(Workbook workbook) {
        this.workbook = workbook;
    }

    public static SheetIterable sheetsOf(Workbook workbook) {
        return new SheetIterable(workbook);
    }

    @Override
    public Iterator<Sheet> iterator() {
        return new SheetIterator(workbook);
    }

}
