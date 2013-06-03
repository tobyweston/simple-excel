package bad.robot.excel.style;

import bad.robot.excel.AbstractValueType;

public class ForegroundColour extends AbstractValueType<Colour> {

    public static ForegroundColour foregroundColour(Colour value) {
        return new ForegroundColour(value);
    }

    private ForegroundColour(Colour value) {
        super(value);
    }
}
