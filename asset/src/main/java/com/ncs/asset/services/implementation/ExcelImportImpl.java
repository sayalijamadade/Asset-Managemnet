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
			  	  String assetId,typeName,groupName,assetName,vendorName,cost,description,location=null;
			  	  
		          Class.forName("com.mysql.jdbc.Driver"); 
				  Connection con = (Connection)
				  DriverManager.getConnection("jdbc:mysql://10.176.45.187:3306/assets","root",
				  "root");
				  con.setAutoCommit(false);
				     PreparedStatement pstm = null; 
		       
		        XSSFWorkbook workbook = (XSSFWorkbook)WorkbookFactory.create(new File("C://Users//P1326048//Documents//asf2.xlsx"));
		        XSSFSheet sheet = workbook.getSheetAt(0);
		        Row row;
		         Statement st=con.createStatement();
		        for(int i=1; i<=sheet.getLastRowNum(); i++){  //points to the starting of excel i.e excel first row
		            row = (Row) sheet.getRow(i);  //sheet number

		            if( row.getCell(0)==null) { assetId = "0";}  //suppose excel cell is empty then its set to 0 the variable
		            else 
		              {assetId = row.getCell(0).toString();   //else copies cell data to SrNo variable
		              System.out.println(assetId);
		              }

		            if( row.getCell(5)==null) { typeName = "0"; }
		            else 
		              {typeName= row.getCell(5).toString();
		              System.out.println(typeName);
		              }

		            if( row.getCell(4)==null) { groupName = "0";   }
		            else  {groupName   = row.getCell(4).toString();
		            System.out.println(groupName);

		            }

		            if( row.getCell(1)==null) { assetName = "0";   }
		            else  {assetName   = row.getCell(1).toString();
		            System.out.println(assetName);
		            }

		            if( row.getCell(6)==null) { vendorName = "0";}  //suppose excel cell is empty then its set to 0 the variable
		            else 
		              {vendorName = row.getCell(6).toString();   //else copies cell data to SrNo variable
		              System.out.println(vendorName);
		              }

		            if( row.getCell(2)==null) { cost = "0"; }
		            else 
		              {cost= row.getCell(2).toString();
		              System.out.println(cost);
		              }

		            if( row.getCell(3)==null) { description = "0";   }
		            else  {description   = row.getCell(3).toString();
		            System.out.println(description);

		            }

		           

		            if( row.getCell(7)==null) { location = "0";   }
		            else  {location   = row.getCell(7).toString();
		            System.out.println(location);
		            }
		            String sq="INSERT INTO asset(asset_id,type_name,group_name,asset_name,vendor_name,cost,description,location_loc_id) "
		                		   + "VALUES('"+assetId+"','"+typeName+"','"+groupName+"','"+assetName+"','"+vendorName+"','"+cost+"','"+description+"','"+location+"')";               
				
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
