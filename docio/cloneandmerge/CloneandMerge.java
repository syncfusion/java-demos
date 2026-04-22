package cloneandmerge;

import java.io.*;
import com.syncfusion.docio.*;

public class CloneandMerge {
	public static void main(String[] args) throws Exception {
		// Create a new document
		WordDocument doc = new WordDocument();
		// Import the document
		doc.importContent(new WordDocument(new FileInputStream(getDataDir("Northwind.docx"))));
		doc.importContent(new WordDocument(new FileInputStream(getDataDir("Adventure.docx"))),
				ImportOptions.KeepSourceFormatting);
		// Save and close the document
		doc.save("Clone and Merge.docx");
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
		if(!(dir.toString().endsWith("samples")))
			dir = dir.getParentFile();
        dir = new File(dir, "resources");
        dir = new File(dir, path);
        if (dir.isDirectory() == false)
            dir.mkdir();
        return dir.toString();
    }
}
