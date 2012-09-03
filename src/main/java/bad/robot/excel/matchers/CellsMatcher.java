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
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.PoiToExcelCoercions.asExcelRow;

/**
 * Assert the number of cells in two workbooks are the same.
 */
public class CellsMatcher extends TypeSafeDiagnosingMatcher<Row> {

    private final Row expected;

    public static CellsMatcher hasSameNumberOfCellsAs(Row expected) {
        return new CellsMatcher(expected);
    }

    private CellsMatcher(Row expected) {
        this.expected = expected;
    }

    @Override
    protected boolean matchesSafely(Row actual, Description mismatch) {
        mismatch.appendText("got ").appendValue(numberOfCellsIn(actual)).appendText(" cell(s) on row ").appendValue(asExcelRow(expected));
        return expected.getLastCellNum() == actual.getLastCellNum();
    }

    @Override
    public void describeTo(Description description) {
        description.appendValue(numberOfCellsIn(expected)).appendText(" cell(s) on row ").appendValue(asExcelRow(expected));
    }

    /** POI is zero-based */
    private static int numberOfCellsIn(Row row) {
        return row.getLastCellNum();
    }
}
