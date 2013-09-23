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

import bad.robot.excel.cell.*;
import bad.robot.excel.column.ColumnIndex;
import bad.robot.excel.style.NoStyle;
import bad.robot.excel.style.Style;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class DefaultStyledRowBuilder implements RowBuilder, StyledRowBuilder {

    private final Map<ColumnIndex, Cell> cells = new HashMap<ColumnIndex, Cell>();
    private final Style defaultStyle;

    public static StyledRowBuilder aRowWithDefaultStyle(Style defaultStyle) {
        return new DefaultStyledRowBuilder(defaultStyle);
    }

    public static StyledRowBuilder aRow() {
        return new DefaultStyledRowBuilder(new NoStyle());
    }

    private DefaultStyledRowBuilder(Style defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    @Override
    public RowBuilder withBlank(ColumnIndex index) {
        this.cells.put(index, new BlankCell(defaultStyle));
        return this;
    }

    @Override
    public StyledRowBuilder withBlank(ColumnIndex index, Style style) {
        this.cells.put(index, new BlankCell(style));
        return this;
    }

    @Override
    public RowBuilder withString(ColumnIndex index, String text) {
        this.cells.put(index, new StringCell(text, defaultStyle));
        return this;
    }

    @Override
    public StyledRowBuilder withString(ColumnIndex index, String text, Style style) {
        this.cells.put(index, new StringCell(text, style));
        return this;
    }

    @Override
    public RowBuilder withDouble(ColumnIndex index, Double value) {
        this.cells.put(index, new DoubleCell(value, defaultStyle));
        return this;
    }

    @Override
    public StyledRowBuilder withDouble(ColumnIndex index, Double value, Style style) {
        this.cells.put(index, new DoubleCell(value, style));
        return this;
    }

    @Override
    public RowBuilder withInteger(ColumnIndex index, Integer value) {
        this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), defaultStyle));
        return this;
    }

    @Override
    public StyledRowBuilder withInteger(ColumnIndex index, Integer value, Style style) {
        this.cells.put(index, new DoubleCell(new Double(Integer.toString(value)), style));
        return this;
    }

    @Override
    public RowBuilder withDate(ColumnIndex index, Date date) {
        this.cells.put(index, new DateCell(date, defaultStyle));
        return this;
    }

    @Override
    public StyledRowBuilder withDate(ColumnIndex index, Date date, Style style) {
        this.cells.put(index, new DateCell(date, style));
        return this;
    }

    @Override
    public RowBuilder withFormula(ColumnIndex index, String formula) {
        this.cells.put(index, new FormulaCell(formula, defaultStyle));
        return this;
    }

    @Override
    public StyledRowBuilder withFormula(ColumnIndex index, String formula, Style style) {
        this.cells.put(index, new FormulaCell(formula, style));
        return this;
    }

    @Override
    public Row build() {
        return new Row(cells);
    }
}
