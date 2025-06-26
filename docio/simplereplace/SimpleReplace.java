package simplereplace;

import java.io.File;

import com.syncfusion.docio.*;

public class SimpleReplace {
	public static void main(String[] args) throws Exception {

         // Loads the template document.
         WordDocument doc = new WordDocument(getDataDir("FindAndReplace.docx"),FormatType.Docx);
         // Replace the text that matches the case.
         doc.replace("Adventure","Adventures",true,false);
         // Saves and closes the Word Document.
         doc.save("SimpleReplace.docx",FormatType.Docx);
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
