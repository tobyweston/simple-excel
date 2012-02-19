package bad.robot;

import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.HashCodeBuilder;
import org.apache.commons.lang3.builder.ToStringBuilder;

import java.io.Serializable;

import static bad.robot.Assertions.assertNotNull;
import static org.apache.commons.lang3.builder.CompareToBuilder.reflectionCompare;
import static org.apache.commons.lang3.builder.ToStringStyle.SHORT_PREFIX_STYLE;

public abstract class AbstractValueType<T extends Serializable> implements ValueType<T>, Serializable, Comparable<AbstractValueType<?>> {

    private final T value;

    public AbstractValueType(T value) {
        assertNotNull(value);
        this.value = value;
    }

    @Override
    public T value() {
        return value;
    }

    @Override
    public int hashCode() {
        return new HashCodeBuilder().append(value()).toHashCode();
    }

    @Override
    public boolean equals(Object other) {
        if (other == null)
            return false;
        if (other == this)
            return true;
        if (this.getClass().equals(other.getClass()))
            return new EqualsBuilder().append(value(), ((AbstractValueType) other).value()).isEquals();
        return false;
    }

    @Override
    public int compareTo(AbstractValueType<?> other) {
        return reflectionCompare(value(), other.value(), true);
    }

    @Override
    public String toString() {
        return new ToStringBuilder(this, SHORT_PREFIX_STYLE).append(value()).toString();
    }

}
