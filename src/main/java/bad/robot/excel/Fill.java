package bad.robot.excel;

public class Fill {

    private ForegroundColour foregroundColour;

    private Fill(ForegroundColour foregroundColour) {
        this.foregroundColour = foregroundColour;
    }

    public static Fill fill(ForegroundColour foregroundColour) {
        return new Fill(foregroundColour);
    }

    public ForegroundColour getForegroundColour() {
        return foregroundColour;
    }

}

