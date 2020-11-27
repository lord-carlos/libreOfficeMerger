import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyAttribute;
import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.beans.XPropertyState;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.Enum;
import com.sun.star.uno.UnoRuntime;

import java.util.HashSet;
import java.util.Optional;
import java.util.Set;

class HTMLTagStylist {
    public void transferStylesFromCurrentDocumentFragmentToInsertedHTML(XTextRange xTextRangeForCurrentDocumentFragment, XTextDocument htmlxTextDocumentToInsert) {
        XPropertySet propertySetToTransferFrom = UnoRuntime.queryInterface(XPropertySet.class, xTextRangeForCurrentDocumentFragment);
        XPropertyState sourcePropertyState = UnoRuntime.queryInterface(XPropertyState.class, xTextRangeForCurrentDocumentFragment);
        Set<String> propertyNamesToTransfer = getPropertyNamesToTransfer(propertySetToTransferFrom, sourcePropertyState);
        XText textToInsert = htmlxTextDocumentToInsert.getText();
        enumerateChildrenAndTransferPropertiesRecursively(propertySetToTransferFrom, propertyNamesToTransfer, textToInsert);
    }

    private Set<String> getPropertyNamesToTransfer(XPropertySet propertySetToTransferFrom, XPropertyState propertyState) {
        Set<String> propertyNamesToTransfer = new HashSet<>();

        XPropertySetInfo propertySetInfoToTransferFrom = propertySetToTransferFrom.getPropertySetInfo();
        for(Property propertyToTransfer : propertySetInfoToTransferFrom.getProperties()) {
            String propertyName = propertyToTransfer.Name;
            if(isTextProperty(propertyName)) {
                propertyNamesToTransfer.add(propertyName);
            }
        }

        return propertyNamesToTransfer;
    }

    private boolean isTextProperty(String propertyName) {
        return propertyName != null && propertyName.startsWith("Char");
    }

    private void enumerateChildrenAndTransferPropertiesRecursively(XPropertySet propertySetToTransferFrom, Set<String> propertyNamesToTransfer, Object element) {
        XEnumerationAccess xEnumerationAccess = UnoRuntime.queryInterface(XEnumerationAccess.class, element);
        if(xEnumerationAccess != null && xEnumerationAccess.hasElements()) {
            XEnumeration enumeration = xEnumerationAccess.createEnumeration();
            while(enumeration.hasMoreElements()) {
                try {
                    Object nextElement = enumeration.nextElement();
                    transferPropertiesToParagraphOrTextPortion(propertySetToTransferFrom, propertyNamesToTransfer, nextElement);
                } catch (WrappedTargetException | NoSuchElementException e) {
                    throw new UnsupportedOperationException();
                }
            }
        }
    }

    private boolean isElementWithChildren(Object element) {
        if(element == null) {
            return false;
        } else {
            XEnumerationAccess xEnumerationAccess = UnoRuntime.queryInterface(XEnumerationAccess.class, element);
            return xEnumerationAccess != null && xEnumerationAccess.hasElements();
        }
    }

    private void transferPropertiesToParagraphOrTextPortion(XPropertySet propertySetToTransferFrom, Set<String> propertyNamesToTransfer, Object paragraphOrTextElement) {
        // For at finde ud af hvilke services/APIs dette element understøtter, kan vi bruge metoden supportsService() og spørge på en konkret service angivet som String.
        XServiceInfo xInfo = UnoRuntime.queryInterface(XServiceInfo.class, paragraphOrTextElement);

        // Vi kan ikke håndtere tabeller...
        if(!xInfo.supportsService("com.sun.star.text.TextTable")) {
            XTextContent textContent = UnoRuntime.queryInterface(XTextContent.class, paragraphOrTextElement);
            if(isElementWithChildren(textContent)) {
                // Der overføres ikke styles til elementer med children (dvs. non-leaf elementer), da det forstyrrer custom styles fra den indsættet HTML-markup
                enumerateChildrenAndTransferPropertiesRecursively(propertySetToTransferFrom, propertyNamesToTransfer, textContent);
            } else if(xInfo.supportsService("com.sun.star.style.CharacterProperties")) {
                // Access the paragraph portion's property set...
                // the properties in this property set are listed in: com.sun.star.style.CharacterProperties
                XPropertySet propertySetToTransferTo = UnoRuntime.queryInterface(XPropertySet.class, xInfo);
                // Propertystate bruges til at finde ud f.eks. om den nuværende værdi af en bestemt property er DEFAULT eller om den er eksplicit sat til noget.
                XPropertyState propertyState = UnoRuntime.queryInterface(XPropertyState.class, paragraphOrTextElement);

                transferPropertiesAndOverwriteDefaultValues(propertySetToTransferFrom, propertyNamesToTransfer, propertySetToTransferTo, propertyState);
            }
        }
    }

    private void transferPropertiesAndOverwriteDefaultValues(XPropertySet propertySetToTransferFrom,
                                                             Set<String> propertyNamesToTransfer,
                                                             XPropertySet propertySetToTransferTo,
                                                             XPropertyState propertyState) {
        XPropertySetInfo propertySetInfoToTransferFrom = propertySetToTransferFrom.getPropertySetInfo();
        for(Property propertyToTransfer : propertySetInfoToTransferFrom.getProperties()) {
            try {
                String propertyName = propertyToTransfer.Name;
                Object propertyValueToTransfer = propertySetToTransferFrom.getPropertyValue(propertyName);
                // Vi overskriver default values fordi de kommer fra master template, og skal derfor erstattes med styles fra området i skabelonen, hvor indholdet skal indsættes.
                if(propertyValueToTransfer != null &&
                   propertyNamesToTransfer.contains(propertyName) &&
                   !isPropertyReadOnly(propertyToTransfer.Attributes) &&
                   isPropertyValueDefault(propertyState, propertyName)) {
                    propertySetToTransferTo.setPropertyValue(propertyName, propertyValueToTransfer);
                }
            } catch (IllegalArgumentException | PropertyVetoException | UnknownPropertyException ignored) {
                // IGNORED: IllegalArgumentException: ugyldig værdi
                // IGNORED: PropertyVetoException: read-only property eller rejected pga. noget andet
                // IGNORED: UnknownPropertyException: denne property kan vi åbenbart ikke overføre
                debug("Exception ignored while transferring property to inserted HTML content ", ignored);
            } catch (WrappedTargetException e) {
                throw new UnsupportedOperationException();
            }
        }
    }

    private boolean isPropertyReadOnly(short propertyAttributes) {
        return (propertyAttributes & PropertyAttribute.READONLY) != 0;
    }

    private boolean isPropertyValueDefault(XPropertyState propertyState, String propertyName) throws UnknownPropertyException {
        return propertyState != null && Optional.ofNullable(propertyState.getPropertyState(propertyName)).map(Enum::getValue).orElse(-1) == PropertyState.DEFAULT_VALUE.getValue();
    }

    private void debug(String msg, Throwable t) {
        System.out.println(msg);
    }
}
