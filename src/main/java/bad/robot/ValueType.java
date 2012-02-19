package bad.robot;

import java.io.Serializable;

public interface ValueType<T> extends Serializable {

    T value();

}
