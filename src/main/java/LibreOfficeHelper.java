import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.Exception;
import com.sun.star.uno.XComponentContext;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public enum LibreOfficeHelper {
    INSTANCE;
    public static final String LIBRE_OFFICE_DIR = "C:\\Program Files\\LibreOffice\\program\\";
    private XComponentContext ctx;

    LibreOfficeHelper() {
        setup();
    }

    public XTextDocument readDocumentFromDisk(String parthToOdtFile) throws IOException, com.sun.star.io.IOException {

        byte[] bytes = Files.readAllBytes(Paths.get(parthToOdtFile));
        OOInputStream input = new OOInputStream(bytes);
        PropertyValue[] propertyValues = Props.makeProps("InputStream", input, "Hidden", true);

        XComponentLoader xComponentLoader = HTMLTagReplacer.queryUnoObject(XComponentLoader.class, getDesktop());
        XComponent x = xComponentLoader.loadComponentFromURL("private:stream", "_blank", 0, propertyValues);
        XTextDocument xTextDocument = HTMLTagReplacer.queryUnoObject(XTextDocument.class, x);
        return xTextDocument;
    }

    /**
     * Converts the provided document to a byte array.
     *
     * @param xTextDocument is the document to be converted.
     * @return the converted byte array
     */
    public byte[] convertXTextDocumentToByteArray(XComponent xTextDocument) {
        OOOutputStream outputStream = new OOOutputStream();
        PropertyValue[] propertyValues = Props.makeProps("OutputStream", outputStream);
        XStorable xstorable = HTMLTagReplacer.queryUnoObject(XStorable.class, xTextDocument);
        try {
            xstorable.storeToURL("private:stream", propertyValues);
        } catch (com.sun.star.io.IOException e) {
            e.printStackTrace();
        }

        return outputStream.toByteArray();
    }

    /**
     * Returns the integration point to Open Office. This is the desktop instance.
     *
     * @return the desktop instance.
     */
    public Object getDesktop() {
        XMultiComponentFactory mxRemoteServiceManager = ctx.getServiceManager();
        Object desktop = null;
        try {
            desktop = mxRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return desktop;
    }

    private void setup() {
        try {
            System.out.println("##########################################");
            System.out.println("BOOTSTRAPPING FROM THIS DIRECTORY:");
            System.out.println(LIBRE_OFFICE_DIR);
            System.out.println();
            System.out.println("##########################################");
            this.ctx = SocketBootstrap.getDefault().bootstrap(LIBRE_OFFICE_DIR);
        } catch (java.lang.Exception e) {
            throw new UnsupportedOperationException();
        }
    }
}
