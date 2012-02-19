package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;
import bad.robot.excel.BorderStyle;

public class LeftBorder extends AbstractValueType<BorderStyle> {

    private LeftBorder(BorderStyle value) {
        super(value);
    }

    public static LeftBorder left(BorderStyle value) {
        return new LeftBorder(value);
    }
}
