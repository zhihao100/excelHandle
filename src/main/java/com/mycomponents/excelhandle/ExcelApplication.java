package com.mycomponents.excelhandle;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelApplication.class, args);
		// SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
		// Date currentTime = new Date();
		// Calendar calendar = Calendar.getInstance();
		// calendar.setTime(currentTime);
		// calendar.add(calendar.DATE, -1);
		// currentTime = calendar.getTime();
		// System.out.println(formatter.format(currentTime));
		// Calendar calendarse = Calendar.getInstance();
		// calendarse.set(Calendar.DATE, 1); // 设置为该月第一天
		// calendarse.add(Calendar.DATE, -1); // 再减一天即为上个月最后一天
		// Date myDate = calendarse.getTime();
		// System.out.println(formatter.format(myDate));
	}

}
