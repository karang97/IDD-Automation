package test_cases;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class powershell {
	public static void main(String[] args) throws IOException, InterruptedException {
		String scriptPath = "C:\\softwares and jars\\IDD_Automation\\test files\\AdjRule.txt";	
		ProcessBuilder processBuilder=new ProcessBuilder("powershell.exe","-ExecutionPolicy","Bypass","File", scriptPath);
		Process process=processBuilder.start();
		   // Read the output
        BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
        String line;
        while ((line = reader.readLine()) != null) {
            System.out.println(line);  // Output the result
        }
        int exitCode= process.waitFor();
        System.out.println("Exited"+ exitCode);
		
}
}