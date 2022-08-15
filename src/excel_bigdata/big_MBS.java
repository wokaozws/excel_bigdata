

package excel_bigdata;


import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.Format;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class big_MBS {


	public static void main(String[] args) throws IOException, WriteException
	   {

	      File xlsFile = new File("jxl_mbs.xls");
	      // 创建一个工作簿
	      WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
	      // 创建一个工作表
	      WritableSheet sheet = workbook.createSheet("sheet1", 0);
	      int dept =1 ;
	      int user = 0;
	      int dept_son=1;
	      int dept_grand_son=1;
	      int dept_grand_son2=1;
	      int dept_grand_son3=1;
	      int dept_sum=0;//验证部门数用
	  	Format f1 = new DecimalFormat("000000");
	  	Format phone_num = new DecimalFormat("18900000000");
	  	Format excelfont = new DecimalFormat("00");

		      for (int row = 0; row < 50000; row++)
		      {
		         for (int col = 0; col <= 14; col++)
		         {
		        	 if (col==0 )
		        	 {
		        		 sheet.addCell(new Label(col, row, "南京指掌易/" + "1级部门"+dept+"/2级部门"+dept_son+"/3级部门"+dept_grand_son+"/4级部门"+dept_grand_son2+"/5级部门"+dept_grand_son3));
		        		
		        	 }
		        
		        	 if (col==1 )
		        	 {
		        		 user++;
		        			 sheet.addCell(new Label(col, row, "test"+f1.format(user)));

		        			if(user%10000==0)
		        			{
		        				dept++;
		        				dept_sum++;
		        			}
		        			if(user%5000==0)
		        			{
		        				dept_son++;
		        				dept_grand_son=0;
		        				dept_sum++;
		        			}
		        			if(user%5==0)
		        			{
		        				dept_grand_son++;
		        				dept_grand_son2=0;
		        				dept_sum++;
		        			}
		        			if(user%2==0)
		        			{
		        				dept_grand_son2++;
		        				dept_grand_son3=0;
		        				dept_sum++;
		        			}
		        			if(user%1==0)
		        			{	
		        				dept_grand_son3++;
		        				dept_sum++;
		        			}
		        	
		        	    	
		        	 }
		        	 if (col==2 ) {
		        		 sheet.addCell(new Label(col, row, "12345678"));
		        		 
		        	 }
		        	 if (col==3 ) {
		        		 sheet.addCell(new Label(col, row, "大数据性能"+f1.format(user)));
		        	 }
		        	 if (col==4 ) {
		        		 sheet.addCell(new Label(col, row, "999"));
		        	 }
		        	 if (col==5 ||col==6) {
		        		 sheet.addCell(new Label(col, row, phone_num.format(user)));
		        	 }
		        	 if (col==13) {
		        		 sheet.addCell(new Label(col, row, "关闭"));
		        	 }
		        	 System.out.println("当前用户:"+"test"+f1.format(user)+" EXCEL"+excelfont.format(col) +"列完成！"+"  当前的部门数量："+dept_sum);
		        	 
		            // 向工作表中添加数据
		           // sheet.addCell(new Label(col, row, "data" + row + col));
		         }
		      }
		
	      workbook.write();
	      workbook.close();

	      
	      
	      
	   }

	
}


