/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
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
import org.hamcrest.Matcher;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import java.util.ArrayList;
import java.util.List;

import static bad.robot.excel.matchers.CompositeMatcher.allOf;
import static bad.robot.excel.matchers.RowInSheetMatcher.hasSameRow;

public class RowsMatcher extends TypeSafeDiagnosingMatcher<Sheet> {

    private final Sheet expected;
    private final List<Matcher<Sheet>> rowsOnSheet;

    public static RowsMatcher hasSameRowsAs(Sheet expected) {
        return new RowsMatcher(expected);
    }

    private RowsMatcher(Sheet expected) {
        this.expected = expected;
        this.rowsOnSheet = createRowMatchers(expected);
    }

    @Override
    protected boolean matchesSafely(Sheet actual, Description mismatch) {
        return allOf(rowsOnSheet).matchesSafely(actual, mismatch);
    }

    @Override
    public void describeTo(Description description) {
        description.appendText("equality on all rows in ").appendValue(expected.getSheetName());
    }

    private static List<Matcher<Sheet>> createRowMatchers(Sheet sheet) {
        List<Matcher<Sheet>> matchers = new ArrayList<Matcher<Sheet>>();
        for (Row expected : sheet)
            matchers.add(hasSameRow(expected));
        return matchers;
    }

}
