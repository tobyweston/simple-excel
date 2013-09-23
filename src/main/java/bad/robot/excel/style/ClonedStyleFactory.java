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

package bad.robot.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class ClonedStyleFactory implements CellStyleFactory {

    private final org.apache.poi.ss.usermodel.Cell source;

    private ClonedStyleFactory(org.apache.poi.ss.usermodel.Cell source) {
        this.source = source;
    }

    public static ClonedStyleFactory newStyleBasedOn(org.apache.poi.ss.usermodel.Cell source) {
        return new ClonedStyleFactory(source);
    }

    @Override
    public CellStyle create(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(source.getCellStyle());
        return style;
    }
}
