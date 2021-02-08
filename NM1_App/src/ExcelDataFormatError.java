import javax.swing.*;

public class ExcelDataFormatError extends Exception {
	public ExcelDataFormatError(String message, JFrame frame){
		super(message);
		JOptionPane.showMessageDialog(frame, message);
	}
}
