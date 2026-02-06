package wordmltoword;

import java.io.File;

import com.syncfusion.docio.*;

public class WordMLToWord {
	public static void main(String[] args) throws Exception {
		// Loads the template document.
		WordDocument wordDoc = new WordDocument(getDataDir("WordMLtoDoc.xml"),FormatType.WordML);
		// Saves and closes the document.
		wordDoc.save("WordML to Word.docx",FormatType.Docx);
		wordDoc.close();
		System.out.println("Word document generated successfully.");
	}
	/**
     * Get the file path
     *
     * @param path specifies the file path
     */
	public static String getDataDir(String path) {
        File dir = new File(System.getProperty("user.dir"));
		if(!(dir.toString().endsWith("samples")))
			dir = dir.getParentFile();
        dir = new File(dir, "resources");
        dir = new File(dir, path);
        if (dir.isDirectory() == false)
            dir.mkdir();
        return dir.toString();
    }
}
