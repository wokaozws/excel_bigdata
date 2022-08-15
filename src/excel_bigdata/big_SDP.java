

package excel_bigdata;


import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.Format;
import java.util.Scanner;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import static java.lang.Thread.sleep;
//import excel_bigdata.big_MBS;

public class big_SDP {
	public static void main(String[] args) throws IOException, WriteException
	   {
		   String U_Choice;
		   Scanner U_Choice_s= new Scanner(System.in);
		   System.out.print("输入对应数字代表生成用户类型：\n" +
				   "        * 1.MBS用户\n" +
				   "        * 2.SDP用户 \n"  );
		   U_Choice = U_Choice_s.nextLine();

			   if ("1".equals(U_Choice)) {
				   System.out.print("没写！");

			   } else if ("2".equals(U_Choice)) {
				   SDP_UserCreat();

			   } else {
				   System.out.print("输入错误！");

			   }

	   }

	static void SDP_UserCreat() throws IOException, WriteException
	{
		int user_num;
		Scanner user_num_s= new Scanner(System.in);
		System.out.print("你当前选择的是SDP生成模式\n"+"输入想要生成的SDP用户数量：");
		user_num = user_num_s.nextInt();

		int File_num= (int) Math.ceil((user_num-1)/50000)+1;

		try {
			System.out.println("注意！将生成的文件数量:"+File_num+"个");
			sleep(1000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		int dept =1 ;
		int user =0;  //初始用户数字
		int dept_son=1;
		int dept_grand_son=1;
		int dept_grand_son2=1;
		int dept_grand_son3=1;
		int dept_sum=5;//验证部门数用
		Format f1 = new DecimalFormat("000000");
		Format phone_num = new DecimalFormat("18900000000");
		Format excelfont = new DecimalFormat("00");
		for(int flag_filenum=0;flag_filenum<File_num;flag_filenum++){
			int flag_rownum=50000;
			if (File_num-flag_filenum<=1)
			{
				flag_rownum=user_num%50000;
				if(flag_rownum==0){flag_rownum=50000;}
			}
			try {
				System.out.print((flag_filenum+1)+"阶段生成,当前内循环次数:"+flag_rownum+"次,生成中:");
				sleep(1000);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
			File xlsFile = new File("jxl_sdp_"+(flag_filenum+1)+".xls");
			// 创建一个工作簿
			WritableWorkbook workbook = Workbook.createWorkbook(xlsFile);
			// 创建一个工作表
			WritableSheet sheet = workbook.createSheet("sdp用户导入", 0);


			for (int row = 0; row < flag_rownum; row++)
			{
				if(row%15000==0){try {
					System.out.print("*");
					sleep(800);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}}

				for (int col = 0; col <= 14; col++)
				{
					if (col==0 )
					{
						sheet.addCell(new Label(col, row, "组织架构/" + "1级部门"+dept+"/2级部门"+dept_son+"/3级部门"+dept_grand_son+"/4级部门"+dept_grand_son2+"/5级部门"+dept_grand_son3));

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
						if(user%2600==0)
						{
							dept_son++;
							dept_grand_son=0;
							dept_sum++;
						}
						if(user%503==0)
						{
							dept_grand_son++;
							dept_grand_son2=0;
							dept_sum++;
						}
						if(user%121==0)
						{
							dept_grand_son2++;
							dept_grand_son3=0;
							dept_sum++;
						}
						if(user%11==0)
						{
							dept_grand_son3++;
							dept_sum++;
						}


					}
					if (col==2 ) {
						sheet.addCell(new Label(col, row, "12345678"));

					}
					if (col==3 ) {
						sheet.addCell(new Label(col, row, "SDP导入性能"+f1.format(user)));
					}
					if (col==4 ) {
						sheet.addCell(new Label(col, row, "999"));
					}
					if (col==5) {
						sheet.addCell(new Label(col, row, phone_num.format(user)));
					}
					if (col==6) {
						sheet.addCell(new Label(col, row, "test"+f1.format(user)+"@zzy.com"));
					}
				//		System.out.println("当前用户:"+"test"+f1.format(user)+" EXCEL"+excelfont.format(col) +"列完成！"+"  当前的部门数量："+dept_sum);

					// 向工作表中添加数据
					// sheet.addCell(new Label(col, row, "data" + row + col));
				}
			}
			System.out.println("\n****************\n"+(flag_filenum+1)+"阶段生成完毕\n"+"当前用户数："+user+" 当前部门数:"+dept_sum);
			System.out.println("生成文件名：jxl_sdp_"+(flag_filenum+1)+".xls"+"\n****************");
			workbook.write();
			workbook.close();
		}

	}








}


