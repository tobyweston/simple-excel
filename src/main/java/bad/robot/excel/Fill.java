package bad.robot.excel;

public class Fill {

    private final ForegroundColour foregroundColour;

    public static Fill fill(ForegroundColour foregroundColour) {
        return new Fill(foregroundColour);
    }

    private Fill(ForegroundColour foregroundColour) {
        this.foregroundColour = foregroundColour;
    }

    public ForegroundColour getForegroundColour() {
        return foregroundColour;
    }

}

