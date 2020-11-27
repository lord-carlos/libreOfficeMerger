
import com.sun.star.io.XOutputStream;

import java.io.ByteArrayOutputStream;

/**
 * Implementation of the common way to interact with Open Office OutputStream.
 *
 */
public class OOOutputStream extends ByteArrayOutputStream implements XOutputStream {
    public OOOutputStream() {
        super(32768);
    }

    public void writeBytes(byte[] values) throws com.sun.star.io.IOException {
        try {
            this.write(values);
        } catch (java.io.IOException e) {
            throw (new com.sun.star.io.IOException(e.getMessage()));
        }
    }

    public void closeOutput() throws com.sun.star.io.IOException {
        try {
            super.flush();
            super.close();
        } catch (java.io.IOException e) {
            throw (new com.sun.star.io.IOException(e.getMessage()));
        }
    }

    @Override
    public void flush() {
        try {
            super.flush();
        } catch (java.io.IOException e) {
        }
    }
}

