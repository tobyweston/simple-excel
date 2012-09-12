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
import org.hamcrest.Description;
import org.hamcrest.Matcher;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.PoiToExcelCoercions.asExcelCoordinate;
import static bad.robot.excel.matchers.CellType.adaptPoi;

public class CellMatcher extends TypeSafeDiagnosingMatcher<Cell> {

    private final bad.robot.excel.Cell expected;

    private CellMatcher(bad.robot.excel.Cell expected) {
        this.expected = expected;
    }

    public static Matcher<Cell> equalTo(bad.robot.excel.Cell expected) {
        return new CellMatcher(expected);
    }

    @Override
    protected boolean matchesSafely(Cell cell, Description mismatch) {
        if (!adaptPoi(cell).equals(expected)) {
            mismatch.appendText("cell at ").appendValue(asExcelCoordinate(cell)).appendText(" contained ").appendValue(adaptPoi(cell)).appendText(" expected ").appendValue(expected);
            return false;
        }
        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendValue(expected);
    }
}
