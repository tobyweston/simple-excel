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

package bad.robot.excel.matchers;

import org.apache.poi.ss.usermodel.Workbook;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

public class SheetNumberMatcher extends TypeSafeDiagnosingMatcher<Workbook> {

    private final Workbook expected;

    private SheetNumberMatcher(Workbook expected) {
        this.expected = expected;
    }

    public static SheetNumberMatcher hasSameNumberOfSheetsAs(Workbook expected) {
        return new SheetNumberMatcher(expected);
    }

    @Override
    protected boolean matchesSafely(Workbook actual, Description mismatch) {
        if (expected.getNumberOfSheets() != actual.getNumberOfSheets()) {
            mismatch.appendText("got " ).appendValue(actual.getNumberOfSheets()).appendText(" sheet(s) expected ").appendValue(expected.getNumberOfSheets());
            return false;
        }
        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendValue(expected.getNumberOfSheets()).appendText(" sheet(s)");
    }
}
