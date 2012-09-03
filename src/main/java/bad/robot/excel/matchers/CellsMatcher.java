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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hamcrest.Description;
import org.hamcrest.Matcher;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import java.util.ArrayList;
import java.util.List;

import static bad.robot.excel.PoiToExcelCoercions.asExcelRow;
import static bad.robot.excel.matchers.CellMatcher.hasSameCell;
import static bad.robot.excel.matchers.CompositeMatcher.allOf;

public class CellsMatcher extends TypeSafeDiagnosingMatcher<Row> {

    private final Row expected;
    private final List<Matcher<Row>> cellsOnRow;

    public static CellsMatcher hasSameCellsAs(Row expected) {
        return new CellsMatcher(expected);
    }

    private CellsMatcher(Row expected) {
        this.expected = expected;
        this.cellsOnRow = createCellMatchers(expected);
    }

    @Override
    protected boolean matchesSafely(Row actual, Description mismatch) {
        return allOf(cellsOnRow).matchesSafely(actual, mismatch);
    }

    @Override
    public void describeTo(Description description) {
        description.appendText("equality of all cells on row ").appendValue(asExcelRow(expected));
    }

    private static List<Matcher<Row>> createCellMatchers(Row row) {
        List<Matcher<Row>> matchers = new ArrayList<Matcher<Row>>();
        for (Cell expected : row)
            matchers.add(hasSameCell(expected));
        return matchers;
    }

}
