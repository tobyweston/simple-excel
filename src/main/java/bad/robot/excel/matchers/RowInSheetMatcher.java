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

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.PoiToExcelCoercions.asExcelRow;
import static bad.robot.excel.matchers.CellNumberMatcher.hasSameNumberOfCellsAs;
import static bad.robot.excel.matchers.CellsMatcher.hasSameCellsAs;
import static bad.robot.excel.matchers.RowMissingMatcher.rowIsPresent;

public class RowInSheetMatcher extends TypeSafeDiagnosingMatcher<Sheet> {

    private final Row expected;
    private int rowIndex;

    public static RowInSheetMatcher hasSameRow(Row expected) {
        return new RowInSheetMatcher(expected);
    }

    private RowInSheetMatcher(Row expected) {
        this.expected = expected;
        this.rowIndex = expected.getRowNum();
    }

    @Override
    protected boolean matchesSafely(Sheet actualSheet, Description mismatch) {
        Row actual = actualSheet.getRow(rowIndex);

        if (!rowIsPresent(expected).matchesSafely(actual, mismatch))
            return false;

        if (!hasSameNumberOfCellsAs(expected).matchesSafely(actual, mismatch))
            return false;

        if (!hasSameCellsAs(expected).matchesSafely(actual, mismatch))
            return false;

        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendText("equality of row ").appendValue(asExcelRow(expected));
    }


}
