import com.sun.star.beans.PropertyValue;
import com.sun.star.datatransfer.UnsupportedFlavorException;
import com.sun.star.datatransfer.XTransferableSupplier;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XSearchDescriptor;
import com.sun.star.util.XSearchable;
import com.sun.star.view.XSelectionSupplier;

/**
 * Find a Tag in an office document and replaces it with HTML
 */
class HTMLTagReplacer {

    private final HTMLTagStylist htmlTagStylist;
    static final String TAG_PREFIX = "[";
    static final String TAG_POSTFIX = "]";
    public static final String TAG_TO_FIND = "Tag";
    public static final String REPLACEMENT_TEXT = "<html>\n" +
                                              "  <head><meta charset=\"iso-8859-1\"/>\n" +
                                              "\n" +
                                              "  <style type=\"text/css\">\n" +
                                              "BASE_SIZE 11</style>\n" +
                                              "</head>\n" +
                                              "  <body>\n" +
                                              "    <ul style=\"margin-top:0px;margin-bottom:0px;\">\n" +
                                              "      <li>\n" +
                                              "first bullet point\n" +
                                              "\n" +
                                              "      </li>\n" +
                                              "      <li>\n" +
                                              "second bullet point\n" +
                                              "\n" +
                                              "      </li>\n" +
                                              "      <li>\n" +
                                              "3rd bullet point\n" +
                                              "\n" +
                                              "      </li>\n" +
                                              "    </ul>\n" +
                                              "</body>\n" +
                                              "</html>";

    HTMLTagReplacer() {
        this.htmlTagStylist = new HTMLTagStylist();
    }

    public XTextDocument replaceHTMLTag(XTextDocument baseFile) {
        // Set properties for at åbne et HTML dokument
        OOInputStream input = new OOInputStream(REPLACEMENT_TEXT.getBytes());
        PropertyValue[] conversionProperties = Props.makeProps("InputStream", input, "Hidden", true, "FilterName", "HTML (StarWriter)");

        XComponentLoader xComponentLoader = queryUnoObject(XComponentLoader.class, LibreOfficeHelper.INSTANCE.getDesktop());

        XComponent htmlDoc = null;
        try {
            // Build a new document based on the HTML
            htmlDoc = xComponentLoader.loadComponentFromURL("private:stream", "_blank", 0, conversionProperties);
            XTextDocument htmlxTextDocument = queryUnoObject(XTextDocument.class, htmlDoc);
            baseFile.getText();
            XTextCursor xOriginalTextCursor;

            // Find tag and select it
            XSearchable xSearchable = queryUnoObject(XSearchable.class, baseFile);
            XSearchDescriptor xSearchDescriptor = xSearchable.createSearchDescriptor();

            xSearchDescriptor.setSearchString(TAG_PREFIX + TAG_TO_FIND + TAG_POSTFIX);
            XTextRange xTextRange = UnoRuntime.queryInterface(XTextRange.class, xSearchable.findFirst(xSearchDescriptor));

            while(xTextRange != null) {
                XText insert = xTextRange.getText();
                xOriginalTextCursor = insert.createTextCursor();
                xOriginalTextCursor.gotoRange(xTextRange, false);

                /*
                  Make sure the styles (font, font size etc.) to be transferred to the base document
                 */
                htmlTagStylist.transferStylesFromCurrentDocumentFragmentToInsertedHTML(xTextRange, htmlxTextDocument);

                //Magic ahead.
                XController htmlCurrentController = htmlxTextDocument.getCurrentController();
                XTransferableSupplier htmlxTransferableSupplier = UnoRuntime.queryInterface(XTransferableSupplier.class, htmlCurrentController);
                XTransferableSupplier xTransferableSupplier = UnoRuntime.queryInterface(XTransferableSupplier.class, baseFile.getCurrentController());
                XSelectionSupplier xOriginalSelectionSupplier = UnoRuntime.queryInterface(XSelectionSupplier.class, baseFile.getCurrentController());
                xOriginalSelectionSupplier.select(xOriginalTextCursor);
                XSelectionSupplier xSelectionSupplier = UnoRuntime.queryInterface(XSelectionSupplier.class, htmlCurrentController);
                XTextCursor xTextCursor = htmlxTextDocument.getText().createTextCursor();
                xTextCursor.gotoStart(false);
                xTextCursor.gotoEnd(true);
                xSelectionSupplier.select(xTextCursor);
                xTransferableSupplier.insertTransferable(htmlxTransferableSupplier.getTransferable());

                xTextRange = UnoRuntime.queryInterface(XTextRange.class, xSearchable.findFirst(xSearchDescriptor));
            }
        } catch (IOException | IllegalArgumentException | UnsupportedFlavorException e) {
            throw new UnsupportedOperationException();
        } finally {
            if(htmlDoc != null) {
                htmlDoc.dispose();
            }
        }

        return baseFile;
    }

    // Found in older Project, no sure why we can't call queryInterface directly
    public static <T> T queryUnoObject(Class<T> clazz, Object instance) {
        int secondsToWait = 3;
        System.out.println("queryUnoObject(). Class: " + clazz + ". Instance: " + instance);
        for(int i = 1; i <= secondsToWait; i++) {
            try {
                // Try to get the object, sleep if applicable.
                T result = UnoRuntime.queryInterface(clazz, instance);
                if(result != null) {
                    return result;
                }

                try {
                    Thread.sleep(1000); // Sleeping one second to wait for initalisation.
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            } catch (com.sun.star.lang.DisposedException e) {
                e.printStackTrace();
                // Timeouted
                if(i == secondsToWait) {
                    throw new UnsupportedOperationException();
                }
            }
        }

        return null;
    }
}
