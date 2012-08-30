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
public class SheetNumberEqualityMatcher extends TypeSafeDiagnosingMatcher<Workbook> {

    private final Workbook expected;

    public static SheetNumberEqualityMatcher hasSameNumberOfSheetsAs(Workbook expected) {
        return new SheetNumberEqualityMatcher(expected);
    }

    private SheetNumberEqualityMatcher(Workbook expected) {
        this.expected = expected;
    }

    @Override
    protected boolean matchesSafely(Workbook actual, Description mismatch) {
        mismatch.appendText("got " ).appendValue(actual.getNumberOfSheets()).appendText(" sheet(s)");
        return expected.getNumberOfSheets() == actual.getNumberOfSheets();
    }

    @Override
    public void describeTo(Description description) {
        description.appendValue(expected.getNumberOfSheets()).appendText(" sheet(s)");
    }
}
