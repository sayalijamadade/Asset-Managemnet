package com.ncs.asset.services.implementation;


import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;

import javax.persistence.EntityManager;
import javax.persistence.Query;

import org.apache.poi.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Session;
import org.springframework.beans.factory.annotation.Autowired;

public class ExcelImportImpl {
	
	@Autowired
	
	public static EntityManager entityManager;
	
	public static Session getSession()
	{
		return entityManager.unwrap(Session.class);
	}
	
	public static void fetchData2() throws EncryptedDocumentException, InvalidFormatException {
		  
		try{
			  	  String asset_group,asset_no,capitalisation_date,department,description1,description2,employee_code,employee_name,reason,serial_number,status,vendor,vendor_name,location_loc_id=null;
			  	  
		          Class.forName("com.mysql.jdbc.Driver"); 
				  Connection con = (Connection)
				  DriverManager.getConnection("jdbc:mysql://10.176.45.187:3306/assets","root",
				  "root");
				  con.setAutoCommit(false);
				     PreparedStatement pstm = null; 
		       
		        XSSFWorkbook workbook = (XSSFWorkbook)WorkbookFactory.create(new File("C://Users//P1326048//Documents//template.xlsx"));
		        XSSFSheet sheet = workbook.getSheetAt(0);
		        Row row;
		         Statement st=con.createStatement();
		        for(int i=1; i<=sheet.getLastRowNum(); i++){  //points to the starting of excel i.e excel first row
		            row = (Row) sheet.getRow(i);  //sheet number

		            if( row.getCell(0)==null) { asset_group = "0";}  //suppose excel cell is empty then its set to 0 the variable
		            else 
		              {asset_group = row.getCell(0).toString();   //else copies cell data to SrNo variable
		              System.out.println(asset_group);
		              }
		            
		            if( row.getCell(1)==null) { asset_no = "0"; }
		            else 
		              {asset_no= row.getCell(1).toString();
		              System.out.println(asset_no);
		              }

		            if( row.getCell(2)==null) { description1 = "0";   }
		            else  {description1   = row.getCell(2).toString();
		            System.out.println(description1);
		            }

		            if( row.getCell(3)==null) { description2 = "0";}  //suppose excel cell is empty then its set to 0 the variable
		            else 
		              {description2 = row.getCell(3).toString();   //else copies cell data to SrNo variable
		              System.out.println(description2);
		              }

		            if( row.getCell(4)==null) { employee_name = "0"; }
		            else 
		              {employee_name= row.getCell(4).toString();
		              System.out.println(employee_name);
		              }

		            if( row.getCell(5)==null) { serial_number = "0";   }
		            else  {serial_number   = row.getCell(5).toString();
		            System.out.println(serial_number);

		            }

		           

		            if( row.getCell(6)==null) { vendor_name = "0";   }
		            else  {vendor_name   = row.getCell(6).toString();
		            System.out.println(vendor_name);
		            }
		            
		            if( row.getCell(7)==null) { capitalisation_date = "0";}  //suppose excel cell is empty then its set to 0 the variable
		            else 
		              {capitalisation_date = row.getCell(7).toString();   //else copies cell data to SrNo variable
		              System.out.println(capitalisation_date);
		              }
		            
		            if( row.getCell(8)==null) { employee_code = "0"; }
		            else 
		              {employee_code= row.getCell(8).toString();
		              System.out.println(employee_code);
		              }

		            if( row.getCell(9)==null) { department = "0";   }
		            else  {department   = row.getCell(9).toString();
		            System.out.println(department);

		            }

		            String sq="INSERT INTO asset(asset_group,asset_no,description1,description2,employee_name,serial_number,vendor_name,capitalisation_date,employee_code,department) "
		                		   + "VALUES('"+asset_group+"','"+asset_no+"','"+description1+"','"+description2+
		                		   "','"+employee_name+"','"+serial_number+"','"+vendor_name+
		                		   "','"+capitalisation_date+"','"+employee_code+"','"+department+"')";
				/*
				 * Query query =
				 * getSession().createQuery("insert into asset(stock_code, stock_name)" +
				 * "select stock_code, stock_name from backup_stock"); int result =
				 * query.executeUpdate();
				 */
				 
						
						  pstm = (PreparedStatement) con.prepareStatement(sq);//here we are usingprepared statement because we are calling this statement for each row
						  pstm.execute(); System.out.println("Import rows "+i);
						 
		        }
					
					  con.commit(); pstm.close(); con.close();
					 
		       
		        System.out.println("Success import excel to mysql table");
		      
				
				  }
				catch(ClassNotFoundException e){ System.out.println(e); }catch(SQLException
				  ex){//ex.printStackTrace(); 
					  System.out.println(ex);
				 
		    }catch(IOException ioe){
		    	System.out.println(ioe);
		    }

		}
	
	public static void main(String[] args) {
		try {
		fetchData2();
		}
		catch(Exception e)
		{
			System.out.println(e);
		}
	}

}
