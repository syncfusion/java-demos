package documentprotection;

import java.io.File;

import com.syncfusion.docio.*;

public class DocumentProtection {
	public static void main(String[] args) throws Exception {
		// Initialize Word document.
		WordDocument document = new WordDocument();
		document = new WordDocument();
		// Open an existing template document.
		document.open(getDataDir("TemplateFormFields.docx"), FormatType.Docx);
		// Enforce protection of the document.
		document.protect(ProtectionType.AllowOnlyFormFields);
		// Save and close the document.
		document.save("Document Protection.docx");
		document.close();
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
