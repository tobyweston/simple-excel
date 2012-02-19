package bad.robot.excel.valuetypes;

public class Border {

    private final TopBorder top;
    private final BottomBorder bottom;
    private final LeftBorder left;
    private final RightBorder right;

    private Border(TopBorder top, RightBorder right, BottomBorder bottom, LeftBorder left) {
        this.top = top;
        this.bottom = bottom;
        this.left = left;
        this.right = right;
    }

    public static Border border(TopBorder top, RightBorder right, BottomBorder bottom, LeftBorder left) {
        return new Border(top, right, bottom, left);
    }

    public BottomBorder getBottom() {
        return bottom;
    }

    public TopBorder getTop() {
        return top;
    }

    public LeftBorder getLeft() {
        return left;
    }

    public RightBorder getRight() {
        return right;
    }
}
