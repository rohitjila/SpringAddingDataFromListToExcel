package com.springcore;

import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

import javax.naming.Name;

import java.util.List;
import java.util.Map;

import static com.springcore.ExcelWriter.writeToExcel;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {
    public static void main(String[] args) {
        ApplicationContext context = new ClassPathXmlApplicationContext("META-INF/config_beans.xml");
        Person p = (Person)context.getBean("employees1");
        String filePath = "C:/Demo/employee_data.xlsx";
        writeToExcel(p.getEmployees(), filePath);
    }
}