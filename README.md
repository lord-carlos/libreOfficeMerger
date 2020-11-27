# LibreOffice HTML Merge test app

### How to run:
* pull code
* gradlew run
* See `src/main/java/resources/` `output.odt` file.

### The "problem"

The `basefile.odt` has a 3 text linies. The middle one is a tag `[Tag]` that we will replace with some HTML text.

Observe:

* The merged HTML code, the bullet points, are on page2

Expected behaviour:

* Merged Bulletpoints should be on page 1


The merge creates a `style:master-page-name="HTML"` for the `style:style` for the first bullet point.

But why?
The important code is in `HTMLTagReplacer.java`