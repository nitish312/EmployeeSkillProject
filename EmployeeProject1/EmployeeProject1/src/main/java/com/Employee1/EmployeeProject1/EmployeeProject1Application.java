package com.Employee1.EmployeeProject1;

import com.Employee1.EmployeeProject1.controller.ExcelController;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.Scanner;

@SpringBootApplication
public class EmployeeProject1Application {

	public static void main(String[] args) {
		SpringApplication.run(EmployeeProject1Application.class, args);

		// Take input from the user
		Scanner scanner = new Scanner(System.in);
		System.out.print("Enter a skill to filter: ");
		String skill = scanner.nextLine();
		scanner.close();

		// Call the ExcelController logic
		ExcelController excelController = new ExcelController();
		String result = excelController.filterAndSaveExcel(skill);
		System.out.println(result);
	}
}
