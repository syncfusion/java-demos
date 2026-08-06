package findandhighlight;

import java.io.File;
import java.util.regex.Pattern;

import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.drawing.ColorSupport;
import com.syncfusion.javahelper.system.text.regularExpressions.MatchSupport;

public class FindandHighlight {
	public static void main(String[] args) throws Exception {

        // Loads the template document.
        WordDocument doc = new WordDocument(getDataDir("FindAndReplace.docx"),FormatType.Docx);
        // Finds the first occurrence of a particular text in the document.
        Pattern regex = Pattern.compile(MatchSupport.trimPattern("Adventure"));
        TextSelection text = doc.find(regex);
        // Sets the color for the found text as single text range.
        text.getAsOneRange().getCharacterFormat().setHighlightColor((ColorSupport.getGreen()).clone());
        // Saves and closes the document.
        doc.save("FindandHighlight.docx");
        doc.close();

        System.out.println("Word document generated successfully.");
	}

	/**
	 * Get the file path
	 * 
	 * @param path specifies the file path
	 */
	public static String getDataDir(String path) {
		File dir = new File(System.getProperty("user.dir"));
		if (!(dir.toString().endsWith("samples")))
			dir = dir.getParentFile();
		dir = new File(dir, "resources");
		dir = new File(dir, path);
		if (dir.isDirectory() == false)
			dir.mkdir();
		return dir.toString();
	}

}
