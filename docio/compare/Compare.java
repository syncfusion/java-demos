package compare;

import java.io.File;

import com.syncfusion.docio.*;

public class Compare {
	public static void main(String[] args) throws Exception {
		//Load the original document.
        WordDocument originalDocument = new WordDocument(getDataDir("OriginalDocument.docx"), FormatType.Docx);
        //Load the revised document.
        WordDocument revisedDocument = new WordDocument(getDataDir("RevisedDocument.docx"), FormatType.Docx);
        //Compare the original document with the revised document.
        originalDocument.compare(revisedDocument, "Nancy Davolio", LocalDateTime.now().minusDays(1));
        //Save the word document.
        originalDocument.save("Sample.docx");
        //Close the word documents.
        originalDocument.close();
        revisedDocument.close();
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
