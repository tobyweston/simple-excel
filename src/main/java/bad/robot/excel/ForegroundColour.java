package bad.robot.excel;

public class ForegroundColour extends AbstractValueType<Colour> {
    private ForegroundColour(Colour value) {
        super(value);
    }

    public static ForegroundColour foregroundColour(Colour value) {
        return new ForegroundColour(value);
    }
}
