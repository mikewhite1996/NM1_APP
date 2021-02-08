import javax.swing.JFrame;
import javax.swing.JOptionPane;

public class FileNameException extends Exception {
	public FileNameException(String message, JFrame frame){
		super(message);
		JOptionPane.showMessageDialog(frame, message);
	}
}
