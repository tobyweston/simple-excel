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

import org.hamcrest.Description;
import org.hamcrest.Matcher;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

class Mismatches<T> {

    private final List<Matcher<T>> mismatches = new ArrayList<Matcher<T>>();

    public void discover(T actual, Iterable<Matcher<T>> matchers) {
        for (Matcher<T> matcher : matchers) {
            if (!matcher.matches(actual))
                mismatches.add(matcher);
        }
    }

    public void describeTo(Description description, T actual) {
        Iterator<Matcher<T>> iterator = mismatches.iterator();
        while (iterator.hasNext()) {
            iterator.next().describeMismatch(actual, description);
            if (iterator.hasNext())
                description.appendText(", ");
        }
    }

    public boolean found() {
        return !mismatches.isEmpty();
    }
}
