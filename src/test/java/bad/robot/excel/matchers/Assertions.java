package bad.robot.excel.matchers;

import org.hamcrest.Matcher;

import static org.hamcrest.MatcherAssert.assertThat;

public class Assertions {

    public static void assertTimezone(Matcher<String> matcher) {
        assertThat("this test is timezone sensitive, you may need to set the user.timezone system property", System.getProperty("user.timezone"), matcher);
    }
}
