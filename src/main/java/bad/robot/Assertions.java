package bad.robot;

public class Assertions {

    public static void assertNotNull(Object object) {
        if (object == null)
            throw new IllegalArgumentException("object can not be null");
    }

}
