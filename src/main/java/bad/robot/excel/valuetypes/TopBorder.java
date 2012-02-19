package bad.robot.excel.valuetypes;

import bad.robot.AbstractValueType;
import bad.robot.excel.BorderStyle;

public class TopBorder extends AbstractValueType<BorderStyle> {

    private TopBorder(BorderStyle value) {
        super(value);
    }

    public static TopBorder top(BorderStyle value) {
        return new TopBorder(value);
    }
}
