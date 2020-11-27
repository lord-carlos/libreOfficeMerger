import com.sun.star.text.XTextDocument;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class Main {
    public static final String RELATIVE_PATH = "src\\main\\resources\\";
    public static final String PATH_TO_OUTPUT_FILE = RELATIVE_PATH + "output.odt";
    public static final String PATH_TO_ODT_FILE = RELATIVE_PATH + "base.odt";

    public static void main(String[] args) throws IOException, com.sun.star.io.IOException {
        // !
        // You might want to adjust the LIBRE_OFFICE_DIR in the LibreOfficeHelper.java class!

        XTextDocument baseDocument = LibreOfficeHelper.INSTANCE.readDocumentFromDisk(PATH_TO_ODT_FILE);

        HTMLTagReplacer program = new HTMLTagReplacer();
        XTextDocument returnDocument = program.replaceHTMLTag(baseDocument);

        Files.write(Paths.get(PATH_TO_OUTPUT_FILE), LibreOfficeHelper.INSTANCE.convertXTextDocumentToByteArray(returnDocument));


        System.out.println("Done! PLEASE KILL THE APP.");
        SocketBootstrap.getDefault().disconnect(baseDocument);
        SocketBootstrap.getDefault().disconnect(returnDocument);
    }
}
