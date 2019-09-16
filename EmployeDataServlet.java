package com.nt.servlet;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.nt.dao.EmployeDaoJdbcPersistance;
import com.nt.dto.Employedto;

public class EmployeDataServlet extends HttpServlet{
	
	public void doGet(HttpServletRequest req,HttpServletResponse res)throws ServletException,
	IOException {
		System.out.println("csp");
		//for excel
		res.setContentType("application/vnd.ms-excel");
		//for download
		res.addHeader("Content-Disposition","attachment;filename=e,ploye.xls"); 
		try {
			//create book
	          HSSFWorkbook book = new HSSFWorkbook();
	          //create sheet
	        HSSFSheet  sheet = book.createSheet("employedetailes");
	      //3. Create multiple
	        HSSFRow r1=sheet.createRow(0);
	        //4. create Cells
	        r1.createCell(0).setCellValue("empid");
	        r1.createCell(1).setCellValue("empnm");
	        r1.createCell(2).setCellValue("empadd");
	        r1.createCell(3).setCellValue("empsalary");
	        
	      //call JDBC
	        List<Employedto> list=EmployeDaoJdbcPersistance.getAllEmployeDetailes();
	        int count=1;
	        for(Employedto e:list) {
	        HSSFRow r=sheet.createRow(count); 
	        r.createCell(0).setCellValue(e.getEMPID());
	        r.createCell(1).setCellValue(e.getEMPNM());
	        r.createCell(2).setCellValue(e.getEMPADD());
	        r.createCell(3).setCellValue(e.getEMPSALARY());
	        count=count+1;
	        }
	        //5. Write to OutputStream 
	        book.write(res.getOutputStream());
	        book.close();
	        
		}catch(Exception e) {
			e.printStackTrace();
		}
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
