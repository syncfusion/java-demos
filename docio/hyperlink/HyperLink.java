package hyperlink;

import java.io.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;

public class HyperLink {
	public static void main(String[] args) throws Exception {
		// Creats a new document.
		WordDocument document = new WordDocument();
		// Adds a new section to the document.
		IWSection section = document.addSection();
		// Adds a new paragraph.
		IWParagraph para = section.addParagraph();
		// Adds a text range.
		para.appendText("Inserting Hyperlink");
		// Applies build-in style to the paragraph.
		para.applyStyle(BuiltinStyle.Title);
		// Adds a new paragraph.
		section.addParagraph();
		// Creates web hyperlink.
		webHyperlink(section);
		// Creates email hyperlink.
		mailHyperlink(section);
		// Creates image hyperlink.
		imageHyperlink(document, section);
		// Creates file hyperlink.
		fileHyperlink(section);
		// Creates bookmark hyperlink.
		bookmarkHyperlink(document, section);
		// Saves and closes the document.
		document.save("Hyperlink.docx");
		document.close();
		System.out.println("Word document generated successfully.");
	}

	/**
	 * Creates web hyperlink.
	 * 
	 * @param section Represents a section to which the hyperlink to be added
	 * 
	 */
	static IWParagraph webHyperlink(IWSection section) throws Exception {
		// Adds paragraph to the section.
		IWParagraph para = section.addParagraph();
		para.appendText("Web Hyperlinks");
		// Applies build-in style to the paragraph.
		para.applyStyle(BuiltinStyle.Heading3);
		// Adds paragraph to the section.
		para = section.addParagraph();
		para.appendText(
				"Hyperlinks to web pages creates a link to a web page, email address or to a program. \nSample Links:");
		// Inserts web hyper link.
		insertHyperLinkToParagraph(section,"Syncfusion", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink, HyperlinkType.WebLink, "http://www.syncfusion.com");
		// Inserts web hyper link.
		insertHyperLinkToParagraph(section,"Google", FieldType.FieldHyperlink, BuiltinStyle.Hyperlink, HyperlinkType.WebLink,"http://www.google.com");
		// Inserts web hyper link.
		insertHyperLinkToParagraph(section,"MSN", FieldType.FieldHyperlink, BuiltinStyle.Hyperlink ,HyperlinkType.WebLink,"http://www.msn.com");		
		return para;
	}

	/**
	 * Creates email hyperlink.
	 * 
	 * @param section Represents a section to which the hyperlink to be added
	 * 
	 */

	static IWParagraph mailHyperlink(IWSection section) throws Exception {
		// Adds paragraph to the section.
		IWParagraph para = section.addParagraph();
		para = section.addParagraph();
		para.appendText("E-mail Hyperlinks");
		// Applies build-in style to the paragraph.
		para.applyStyle(BuiltinStyle.Heading3);
		// Adds paragraph to the section.
		para = section.addParagraph();
		para.appendText("Hyperlinks that links to e-mail.");
		// Inserts mail hyper link.
		insertHyperLinkToParagraph(section,"John", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.EMailLink,"mailto:john@gmail.com"); 
		// Inserts mail hyper link.
		insertHyperLinkToParagraph(section,"Eric", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.EMailLink,"mailto:eric@yahoo.com"); 
		// Inserts mail hyper link.
		insertHyperLinkToParagraph(section,"David", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.EMailLink,"mailto:david@hotmail.com"); 
		return para;
	}

	/**
	 * Creates image hyperlink.
	 * 
	 * @param document Represents the word document instance
	 * @param section  Represents a section to which the hyperlink to be added
	 * 
	 */
	static IWParagraph imageHyperlink(WordDocument document, IWSection section) throws Exception {
		// Adds paragraph to the section.
		IWParagraph para = section.addParagraph();
		para = section.addParagraph();
		para.appendText("Image Hyperlink");
		// Applies build-in style to the paragraph.
		para.applyStyle(BuiltinStyle.Heading3);
		// Adds paragraph to the section.
		para = section.addParagraph();
		para.appendText("Hyperlinks to image creates link to a web page, email address or to a program.");
		para = section.addParagraph();
		WPicture mImage1 = new WPicture(document);
		mImage1.loadImage(new FileInputStream(getDataDir("syncfusion_logo.gif")));
		IWField field = para.appendField("Hyperlink", FieldType.FieldHyperlink);
		Hyperlink hyperLink = new Hyperlink((WField)field);
		hyperLink.setType(HyperlinkType.WebLink);
		hyperLink.setUri("http://www.syncfusion.com");
		hyperLink.setPictureToDisplay(mImage1);
		para = section.addParagraph();
		WPicture mImage2 = new WPicture(document);
		mImage2.loadImage(new FileInputStream(getDataDir("google.png")));
		field = para.appendField("Hyperlink", FieldType.FieldHyperlink);
		hyperLink = new Hyperlink((WField)field);
		hyperLink.setType(HyperlinkType.WebLink);
		hyperLink.setUri("http://www.google.com");
		hyperLink.setPictureToDisplay(mImage2);
		para = section.addParagraph();
		WPicture mImage3 = new WPicture(document);
		mImage3.loadImage(new FileInputStream(getDataDir("yahoo.gif")));
		field = para.appendField("Hyperlink", FieldType.FieldHyperlink);
		hyperLink = new Hyperlink((WField)field);
		hyperLink.setType(HyperlinkType.WebLink);
		hyperLink.setUri("http://www.yahoo.com");
		hyperLink.setPictureToDisplay(mImage3);
		return para;
	}

	/**
	 * Creates file hyperlink.
	 * 
	 * @param section Represents a section to which the hyperlink to be added
	 * 
	 */
	static IWParagraph fileHyperlink(IWSection section) throws Exception {
		IWParagraph para = section.addParagraph();
		para = section.addParagraph();
		para = section.addParagraph();
		para.appendText("File Hyperlinks");
		para.applyStyle(BuiltinStyle.Heading3);
		para = section.addParagraph();
		para.appendText("Hyperlinks to files creates links to other files, an image and so on.");
		// Inserts file hyper link.
		insertHyperLinkToParagraph(section,"File 1", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.FileLink,"src/test/resources/SampleBrowserData/Images/DocIO.gif");
		// Inserts file hyper link.
		insertHyperLinkToParagraph(section,"File 2", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.FileLink,"src/test/resources/SampleBrowserData/Images/XlsIO.gif");
		// Inserts file hyper link.
		insertHyperLinkToParagraph(section,"File 3", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.FileLink,"src/test/resources/SampleBrowserData/data/pdf.gif");
		return para;
	}

	/**
	 * Creates bookmark hyperlink.
	 * 
	 * @param document Represents the word document instance
	 * @param section  Represents a section to which the hyperlink to be added
	 * 
	 */
	static IWParagraph bookmarkHyperlink(WordDocument document, IWSection section) throws Exception {
		section = document.addSection();
		section.setBreakCode(SectionBreakCode.NewPage);
		IWParagraph para = section.addParagraph();
		para.appendBookmarkStart("Introduction");
		para.appendText("What is Hyperlink?").getCharacterFormat().setBold(true);
		para.appendBookmarkEnd("Introduction");
		para.appendText(
				"\nA hyperlink is a reference or navigation element in a document to another section of the same document or to another document that may be on or part of a (different) domain.");
		para = section.addParagraph();
		para = section.addParagraph();
		para.appendBookmarkStart("Insert");
		para.appendText("How to create a Hyperlink?").getCharacterFormat().setBold(true);
		para.appendBookmarkEnd("Insert");
		para.appendText(
				"\nSyncfusion.DocIO.DLS.IWField field = para.AppendField(\"Syncfusion\", Syncfusion.DocIO.FieldType.FieldHyperlink);\n");
		para.appendText("para.ApplyStyle(Syncfusion.DocIO.DLS.BuiltinStyle.Hyperlink);\n");
		para.appendText("Syncfusion.DocIO.DLS.Hyperlink hyperLink = new Hyperlink(field as WField);\n");
		para.appendText("hyperLink.Type = Syncfusion.DocIO.DLS.HyperlinkType.WebLink;\n");
		para.appendText("hyperLink.Uri = \"http://www.syncfusion.com\";\n");
		para = section.addParagraph();
		para.appendBookmarkStart("Edit");
		para.appendText("How to edit Hyperlink?").getCharacterFormat().setBold(true);
		para.appendBookmarkEnd("Edit");
		para.appendText("\nSyncfusion.DocIO.DLS.Hyperlink hlink = new Hyperlink(item as WField);\n");
		para.appendText("if(hlink.Type = Syncfusion.DocIO.DLS.HyperlinkType.WebLink)\n");
		para.appendText("{\n");
		para.appendText("if (hlink.TextToDisplay == \"Syncfusion\")\n");
		para.appendText("{\n");
		para.appendText("hlink.TextToDisplay = \"Syncfusion Support\";\n");
		para.appendText("hlink.Uri = \"http://www.syncfusion.com/support/default.aspx\";\n");
		para.appendText("}\n");
		para.appendText("}\n");
		section = document.getSections().get(0);
		para = section.addParagraph();
		para = section.addParagraph();
		para.appendText("Bookmark Hyperlinks");
		para.applyStyle(BuiltinStyle.Heading3);
		para = section.addParagraph();
		para.appendText(
				"A bookmark is a location or selected text on a document that was marked. One can create a hyperlink to a bookmark.");
		// Inserts bookmark hyper link.
		insertHyperLinkToParagraph(section,"What is Hyperlink?", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.Bookmark,"Introduction");
		// Inserts bookmark hyper link.
		insertHyperLinkToParagraph(section,"How to create a Hyperlink?", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.Bookmark,"Insert");
		// Inserts bookmark hyper link.
		insertHyperLinkToParagraph(section,"How to edit Hyperlink?", FieldType.FieldHyperlink,BuiltinStyle.Hyperlink,HyperlinkType.Bookmark,"Edit");
		return para;
	}
	static void insertHyperLinkToParagraph(IWSection section, String fieldName, FieldType fieldType, BuiltinStyle builtinStyle, HyperlinkType hyperlinkType, String value) throws Exception{
		// Adds paragraph to the section.
		IWParagraph para = section.addParagraph();
		// Append field to the paragraph.
		IWField field = para.appendField(fieldName, fieldType);
		// Applies build-in style to paragraph.
		para.applyStyle(builtinStyle);
		// Create hyperlink.
		Hyperlink hyperLink = new Hyperlink((WField)field);
		// sets hyperlink type. 
		hyperLink.setType(hyperlinkType);
		// Checks the hyperlink type is web link or email link then sets uri.
		if(hyperlinkType == HyperlinkType.WebLink || hyperlinkType == HyperlinkType.EMailLink)
		hyperLink.setUri(value);
		// Checks the hyperlink type is file link then sets filepath.
		if(hyperlinkType == HyperlinkType.FileLink)
		hyperLink.setFilePath(value);
		// Checks the hyperlink type is bookmark then sets filepath.
		if(hyperlinkType == HyperlinkType.Bookmark)
		hyperLink.setBookmarkName(value);
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
