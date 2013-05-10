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

import org.apache.poi.ss.usermodel.Sheet;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

public class RowNumberMatcher extends TypeSafeDiagnosingMatcher<Sheet> {

    private final Sheet expected;

    public static RowNumberMatcher hasSameNumberOfRowAs(Sheet expected) {
        return new RowNumberMatcher(expected);
    }

    private RowNumberMatcher(Sheet expected) {
        this.expected = expected;
    }

    @Override
    protected boolean matchesSafely(Sheet actual, Description mismatch) {
        if (expected.getLastRowNum() != actual.getLastRowNum()) {
            mismatch.appendText("got ")
                    .appendValue(numberOfRowsIn(actual))
                    .appendText(" row(s) in sheet ")
                    .appendValue(actual.getSheetName())
                    .appendText(" expected ")
                    .appendValue(numberOfRowsIn(expected));
            return false;
        }
        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendValue(numberOfRowsIn(expected)).appendText(" row(s) in sheet ").appendValue(expected.getSheetName());
    }

    /* POI is zero-based */
    private static int numberOfRowsIn(Sheet sheet) {
        return sheet.getLastRowNum() + 1;
    }
}
