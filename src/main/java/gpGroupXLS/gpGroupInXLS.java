package gpGroupXLS;

import org.apache.logging.log4j.LogManager;

import gpGroupXLS.group.GroupXLS2;

class GPThread extends Thread {
    @SuppressWarnings("unused")
	private static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(GPThread.class);

    private final String jsonFile;
    private final String xlsInputFile;


    public GPThread(String jsonFile, String xlsInputFile) {
        this.jsonFile = jsonFile;
        this.xlsInputFile = xlsInputFile;
    }

    // Override the run method
    @Override
    public void run() {
		GroupXLS2 grpXLS = new GroupXLS2();
		grpXLS.readXLSBuildGroup(jsonFile, xlsInputFile);

    }
}

public class gpGroupInXLS {
    private static final org.apache.logging.log4j.Logger logger = LogManager.getLogger(gpGroupInXLS.class);

    // Display usage instructions
    private void showUsage() {
        System.out.println("Usage: gpGroupInXLS [jsonfile] <xlsfile>");
    }
	
	// Method to display a rotating spinner with colors
    private static void displaySpinner(int index, long elapsedTime) {
        char[] spinnerChars = {'|', '/', '-', '\\'};
        String[] colors = {
            "\u001B[31m", // Red
            "\u001B[32m", // Green
            "\u001B[33m", // Yellow
            "\u001B[34m"  // Blue
        };
        String resetColor = "\u001B[0m"; // Reset color

        // Select color based on the spinner index
        String color = colors[index % colors.length];
        System.out.print("\r" + color + spinnerChars[index % spinnerChars.length] + resetColor + " " + elapsedTime + " ms");
    }

	public static void main(String[] args) {
		try {
			gpGroupInXLS app = new gpGroupInXLS();

			int parmNo;
			String xlsGroupFile = "", configJSON = "" ;

			// Validate command line arguments
			if (args.length == 0 || args.length > 3) {
				app.showUsage();
				return;
			}

			for (parmNo = 0; parmNo < args.length; parmNo++) {
				String arg = args[parmNo];

				// Check for help flag
				if (arg.equals("-h")) {
					app.showUsage();
					return;
				} 
				
				if (parmNo == 0) configJSON = args[parmNo] ;
				if (parmNo == 1) xlsGroupFile = args[parmNo] ;
			}

			// Start processing in a separate thread
			long startTime = System.currentTimeMillis();
			GPThread processingThread = new GPThread(configJSON, xlsGroupFile);
			processingThread.start();

			// Display spinner while the thread is running
			int spinnerIndex = 0;
			while (processingThread.isAlive()) {
				long currentTime = System.currentTimeMillis();
				long elapsedTime = currentTime - startTime;
				displaySpinner(spinnerIndex++, elapsedTime);
				Thread.sleep(100); // Update every 100 milliseconds
			}

			// Wait for the thread to finish
			try {
				processingThread.join();
			} catch (InterruptedException e) {
				logger.error("Error: {}", e.getMessage());
			}

			// Calculate and display elapsed time
			long elapsedTime = System.currentTimeMillis() - startTime;
			double elapsedSeconds = elapsedTime / 1000.0;
			System.out.print("\r" + " ".repeat(50)); // Clear spinner
			System.out.printf("\rElapsed time: %.3f seconds", elapsedSeconds);			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
