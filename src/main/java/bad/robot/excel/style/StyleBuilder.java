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

import bad.robot.excel.cell.DataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import static bad.robot.excel.style.Colour.Black;
import static bad.robot.excel.style.FontColour.fontColour;

public class StyleBuilder implements Style {

    private DataFormat format;
    private Alignment alignment;
    private FontSize fontSize;
    private Border border;
    private Fill fill;
    private FontColour fontColour = fontColour(Black);

    private StyleBuilder() {
    }

    public static StyleBuilder aStyle() {
        return new StyleBuilder();
    }

    public StyleBuilder with(DataFormat format) {
        this.format = format;
        return this;
    }

    public StyleBuilder with(Alignment alignment) {
        this.alignment = alignment;
        return this;
    }

    public StyleBuilder with(FontSize fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public StyleBuilder with(FontColour fontColour) {
        this.fontColour = fontColour;
        return this;
    }

    public StyleBuilder with(Border border) {
        this.border = border;
        return this;
    }

    public StyleBuilder with(Fill fill) {
        this.fill = fill;
        return this;
    }

    private ReplaceExistingStyle build() {
        return new ReplaceExistingStyle(border, format, alignment, fontSize, fontColour, fill);
    }

    @Override
    public void applyTo(Cell cell, Workbook workbook) {
        build().applyTo(cell, workbook);
    }
}