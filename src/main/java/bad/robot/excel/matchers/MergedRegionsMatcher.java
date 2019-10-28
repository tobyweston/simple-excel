/*
 * Copyright (c) 2012-2019, bad robot (london) ltd.
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

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.hamcrest.Description;
import org.hamcrest.TypeSafeDiagnosingMatcher;

import java.util.stream.Collectors;

public class MergedRegionsMatcher extends TypeSafeDiagnosingMatcher<Sheet> {

    private final Sheet expected;

    public static MergedRegionsMatcher hasSameMergedRegions(Sheet expected) {
        return new MergedRegionsMatcher(expected);
    }

    private MergedRegionsMatcher(Sheet expected) {
        this.expected = expected;
    }

    @Override
    protected boolean matchesSafely(Sheet actual, Description mismatch) {
        if (!expected.getMergedRegions().equals(actual.getMergedRegions())) {
            mismatch.appendText("got ")
                .appendValue(mergedRegion(actual))
                .appendText(" as a merged region(s) in sheet ")
                .appendValue(actual.getSheetName())
                .appendText(" expected ")
                .appendValue(mergedRegion(expected))
            ;
            return false;
        }
        return true;
    }

    @Override
    public void describeTo(Description description) {
        description.appendValue(mergedRegion(expected)).appendText(" as a merged region(s) in sheet ").appendValue(expected.getSheetName());
    }

    private String mergedRegion(Sheet sheet) {
        return sheet.getMergedRegions().stream().map(CellRangeAddress::formatAsString).collect(Collectors.joining());
    }

}
