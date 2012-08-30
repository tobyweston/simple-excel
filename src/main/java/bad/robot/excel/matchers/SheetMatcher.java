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

import org.apache.poi.ss.usermodel.Workbook;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import static bad.robot.excel.matchers.SheetNameMatcher.containsSameNamedSheetsAs;
import static bad.robot.excel.matchers.SheetNumberMatcher.hasSameNumberOfSheetsAs;

/**
 * Assert the number of sheets in two workbooks are the same.
 *
 * Will return something like the following using (old style) {@link org.junit.Assert#assertThat(Object, org.hamcrest.Matcher)}
 * <code>
 *     java.lang.AssertionError:
 *     Expected: '2' sheets
 *          got: &lt;org.apache.poi.hssf.usermodel.HSSFWorkbook@3b6f0be8&gt;
 * </code>
 *
 * and, something like the following for {@link org.hamcrest.MatcherAssert#assertThat(Object, org.hamcrest.Matcher)}
 *
 * <code>
 *     java.lang.AssertionError:
 *     Expected: <2> sheet(s)
 *          but: got <1> sheet(s)
 * </code>
 */
public class SheetMatcher extends TypeSafeDiagnosingMatcher<Workbook> {

    private final SheetNameMatcher names;
    private final SheetNumberMatcher numberOfSheets;

    public static SheetMatcher hasSameSheetsAs(Workbook expected) {
        return new SheetMatcher(expected);
    }

    private SheetMatcher(Workbook expected) {
        this.numberOfSheets = hasSameNumberOfSheetsAs(expected);
        this.names = containsSameNamedSheetsAs(expected);
    }

    @Override
    protected boolean matchesSafely(Workbook actual, Description mismatch) {
        if (!numberOfSheets.matchesSafely(actual, mismatch))
            return false;
        return names.matchesSafely(actual, mismatch);
    }

    @Override
    public void describeTo(Description description) {
        numberOfSheets.describeTo(description);
        names.describeTo(description);
    }

}
