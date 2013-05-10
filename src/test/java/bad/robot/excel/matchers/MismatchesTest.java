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

import org.hamcrest.Matcher;
import org.hamcrest.StringDescription;
import org.jmock.Expectations;
import org.jmock.Mockery;
import org.jmock.integration.junit4.JUnit4Mockery;
import org.junit.Test;

import java.util.List;

import static java.util.Arrays.asList;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.containsString;
import static org.hamcrest.Matchers.not;
import static org.hamcrest.core.Is.is;

public class MismatchesTest {

    private final Mockery context = new JUnit4Mockery();

    private final Mismatches<String> mismatches = new Mismatches<String>();
    private final StringDescription description = new StringDescription();
    private final String actual = "I've been checked for mismatches";
    private final Matcher<String> matcher = context.mock(Matcher.class);
    private final List<Matcher<String>> matchers = asList(matcher);

    @Test
    public void noMismatchesOnInitialisation() {
        assertThat(mismatches.found(), is(false));
    }

    @Test
    public void mismatchesFound() {
        context.checking(new Expectations() {{
            oneOf(matcher).matches(actual); will(returnValue(true));
        }});
        assertThat(mismatches.discover(actual, matchers), is(false));
    }

    @Test
    public void noMismatchesFound() {
        context.checking(new Expectations() {{
            oneOf(matcher).matches(actual); will(returnValue(false));
        }});
        assertThat(mismatches.discover(actual, matchers), is(true));
    }

    @Test
    public void descriptionIsEmptyBeforeLookingForMismatches() {
        mismatches.describeTo(description, "not yet checked for mismatches");
        assertThat(description.toString(), is(""));
    }

    @Test
    public void descriptionAfterLookingForMismatchesAndNoneFound() {
        context.checking(new Expectations() {{
            oneOf(matcher).matches(actual); will(returnValue(true));
        }});

        mismatches.discover(actual, matchers);
        mismatches.describeTo(description, "this should not appear in the description as there were no mismatches");
        assertThat(mismatches.found(), is(false));
        assertThat(description.toString(), is(""));
    }

    @Test
    public void descriptionAfterLookingForMismatchesAndMismatchesWereFound() {
        context.checking(new Expectations() {{
            oneOf(matcher).matches(actual); will(returnValue(false));
            oneOf(matcher).describeMismatch(actual, description);
        }});

        mismatches.discover(actual, matchers);
        mismatches.describeTo(description, actual);
        assertThat(mismatches.found(), is(true));
        assertThat(description.toString(), not(containsString(", ")));
    }

    @Test
    public void descriptionAfterLookingForMismatchesAndMultipleMismatchesWereFound() {
        final Matcher<String> matcher1 = context.mock(Matcher.class, "first");
        final Matcher<String> matcher2 = context.mock(Matcher.class, "second");

        context.checking(new Expectations() {{
            oneOf(matcher1).matches(actual); will(returnValue(false));
            oneOf(matcher2).matches(actual); will(returnValue(false));
            oneOf(matcher1).describeMismatch(actual, description);
            oneOf(matcher2).describeMismatch(actual, description);
        }});

        mismatches.discover(actual, asList(matcher1, matcher2));
        mismatches.describeTo(description, actual);
        assertThat(mismatches.found(), is(true));
        assertThat(description.toString(), containsString(",\n          "));
    }
}
