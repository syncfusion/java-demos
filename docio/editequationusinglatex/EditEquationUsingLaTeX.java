package editequationusinglatex;

import java.io.File;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.*;
import com.syncfusion.javahelper.system.collections.generic.ListSupport;

public class EditEquationUsingLaTeX {

	public static void main(String[] args) throws Exception {
		//Opens template document.
		WordDocument document = new WordDocument(getDataDir("EditEquationLaTeXInput.docx"), FormatType.Docx);
		//Find all the equations in the Word document.
		ListSupport<Entity> entities = document.findAllItemsByProperty(EntityType.Math,null,null);
		//Get the first math equation.
		WMath math = (WMath)ObjectSupport.instanceOf(entities.get(0),WMath.class);
		//Update the modified laTeX code to the equation.
		math.getMathParagraph().setLaTeX("\\frac{d}{dx}\\left( {x}^{2}\\right)=k{x}^{k-1}");
		//Save the word document.
		document.save("Sample.docx");
		//Close the word document.
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
