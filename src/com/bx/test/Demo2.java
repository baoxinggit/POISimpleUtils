package com.bx.test;

import java.util.ArrayList;
import java.util.List;

import com.bx.entity.User;
import com.bx.utils.*;

public class Demo2 {
	public static void main(String[] args) {
		
		String fileName = "影虎表";
		String sheetName = "信息";
		String [] cellName = new String [] {"姓名","密码","编号"};
		String path = "E:/";
		List<User> list = new ArrayList<>();
		list.add(new User("1","2",1));
		list.add(new User("2","2",2));
		list.add(new User("3","3",3));
		list.add(new User("4","4",4));
		list.add(new User("5","5",5));
		ExellExportUtil.writeToFile(fileName,sheetName,cellName,path,User.class,list,true);
	}
}
