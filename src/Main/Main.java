package Main;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.util.ArrayList;

import excelUtil.ExcelUtil;

public class Main {
	
	
	public static void main(String[] args) {
		
		txtToSql(2,3,"test","D:/test.xlsx","D:/sql.sql");
	}

	/**
	 * 
	 * @param startRow excel中开始行号
	 * @param endRow   excel中结束行号
	 * @param tableName  生成sql的表名
	 * @param excelPath excel文件存储路径
	 * @param sqlPath  sql文件存储路径
	 */
	public static void txtToSql(int startRow,int endRow,String tableName,String excelPath,String sqlPath){
		
		startRow--;
		endRow--;
		
		//读取excel文件
		File file = new File(excelPath);
		ArrayList<ArrayList<Object>> result = ExcelUtil.readExcel(file);
		String pk = result.get(startRow).get(0).toString(); //第一个字段为主键
		
		//检验结束是否超出
		int excelRowCount = result.size()-1;
		if(excelRowCount<endRow){
			endRow=excelRowCount;
		}
		
		//sql语句
		StringBuffer sql = new StringBuffer();
		sql.append("CREATE TABLE `"+tableName+"`(\n");
		
		for(int i = startRow ;i <= endRow ;i++){
			//获取一行
			ArrayList<Object> oneRow = result.get(i);
			//一行前5列 ，字段名、描述、类型、长度、是否为空
			String keyWord = oneRow.get(0).toString();
			String des = oneRow.get(1).toString();
			String type = oneRow.get(2).toString();
			String len = oneRow.get(3).toString();
			String isNotNull = oneRow.get(4).toString();
			sql.append("\t");
			sql.append("`"+keyWord+"`");
			sql.append(" "+type);
			if(!len.equals("")){
				sql.append("("+Double.valueOf(len).intValue()+")");
			}
			if(isNotNull.equals("是")){
				sql.append(" "+"NOT NULL");
			}
			if(!des.equals("")){
				sql.append(" "+"COMMENT"+" "+"\'"+des+"\'");
			}
			sql.append(",\n");
		} 
		sql.append("\t"+"PRIMARY KEY (`"+pk+"`)\n)");
		
		//写sql文件
		try {  
            File f = new File(sqlPath);  
            if (!f.exists()) {    
                f.createNewFile();// 不存在则创建  
            }  
  
            BufferedWriter output = new BufferedWriter(new FileWriter(f));  
            output.write(sql.toString());  
            output.close();  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
		System.out.println("successful");
	}
	
}
