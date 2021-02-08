import javax.swing.JProgressBar;
import javax.swing.SwingUtilities;

public class fillProgress {
	public static void update(String message, int i, JProgressBar b){
		 while(i <= 100) {
		      if(i <= 30){
		    	  updateBar(i, b);
		    	  updateMessage(message);
		    	  break;
		      }
		      if(i > 30 && i < 70){
		    	  updateBar(i,b);
		    	  updateMessage(message);
		    	  break;
		      }
		      if(i==100){
		    	  updateBar(i,b);
		    	  updateMessage(message);
		    	  break;
		      }
		  }
	}
	
	
	
	public static void updateBar(int i, JProgressBar b){
		b.setValue(i);
	}
	
	public static String updateMessage(String message){
		return message;
	}
}
