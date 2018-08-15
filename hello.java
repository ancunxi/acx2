package com.offcn.web;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.mysql.fabric.Response;
import com.offcn.bean.NewStudent;
import com.offcn.bean.Phone;
import com.offcn.service.NewStudentInfoService;


@Controller
public class FileUploadController {
	@Autowired
	NewStudentInfoService newStudentInfoService;
	
	
	@RequestMapping(value="/importexcel",method=RequestMethod.POST)
	public String uploadExcel(HttpServletRequest request,Model model,@RequestParam("file") MultipartFile file) throws Exception, InvalidFormatException, IOException{
		List<Phone> list = new ArrayList<Phone>();
		
		//获取服务器端的路径
		String path = request.getServletContext().getRealPath("upload");
		//获得上传文件的文件名
		String fileName = file.getOriginalFilename();
		//创建目标file
		File targetFile = new File(path+"\\"+fileName);
		//创建存储目录
		File targetPath = new File(path);
		//判断服务器目录是否存在，如果不存在创建目录
		if(!targetPath.exists()){
			targetPath.mkdir();
		}
		//把上传的文件存储到服务器
		try {
			file.transferTo(targetFile);
		} catch (IllegalStateException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//读取上传到服务器端的文件
		Workbook workbook = WorkbookFactory.create(targetFile);
		//得到单元格里面的工作表
		Sheet sheet = workbook.getSheet("sheet1");
		//判断行数
		int rownum = sheet.getPhysicalNumberOfRows();
		//遍历行，取出每个单元格
		for(int i=0;i<rownum;i++){
			if(i==0){
				continue;
			}
			Row row = sheet.getRow(i);
			//判断当前行单元格的数量
			int cellnum = row.getPhysicalNumberOfCells();
			StringBuffer buf = new StringBuffer();
			//循环拿到单元格
			for(int j=0;j<cellnum;j++){
				Cell cell = row.getCell(j);
				//判断当前单元格是不是string类型的
				if(cell.getCellType()==HSSFCell.CELL_TYPE_STRING){
					buf.append(cell.getStringCellValue()+"~");
				}else if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
					//创建数字格式化工具类
					DecimalFormat df = new DecimalFormat("####");
					//把单元格读取到的数字，进行格式化防止科学计数法形式显示
					buf.append(df.format(cell.getNumericCellValue())+"~");
				}	
			}
		//将读取的本行内容进行相应格式的转换
			String hang = buf.toString();
			//对本行内容进行拆分
			String[] split = hang.split("~");
			//将信息封装到学生对象中
		/*	NewStudent stu = new NewStudent();
			stu.setName(split[1]);
			stu.setScore(Integer.parseInt(split[2]));
			stu.setPhone(split[3]);*/
			
			
			Phone phone = new Phone();
			phone.setId(Integer.parseInt(split[0]));
			phone.setMobilenumber(Integer.parseInt(split[1]));
			phone.setMobilearea(split[2]);
			phone.setMobiletype(split[3]);
			phone.setAreacode(Integer.parseInt(split[4]));
			phone.setPostcode(Integer.parseInt(split[5]));
			
			
			
			
			
			
			
			list.add(phone);
			/*System.out.println("上传学生信息："+stu);	*/	
		}
		newStudentInfoService.save(list);
		return "success";
	}
	@RequestMapping(value="/downLoadexcel")
	public void downloadexcel(HttpServletRequest request,HttpServletResponse response) throws Exception, IOException{
		List<NewStudent> list = newStudentInfoService.getAllStudent();
		//获取服务器端路径
		String path = request.getServletContext().getRealPath("down");
		String fileName = "testexcel.xlsx";
		//创建存储的文件
		File targetFile = new File(path+"\\"+fileName);
		//创建存储目录
		File targetPath = new File(path);
		//判断服务器端是否存在，如果不存在创建目录
		if(!targetPath.exists()){
			targetPath.mkdir();
		}
		//生成excel
		XSSFWorkbook workbook = new XSSFWorkbook();
		//创建工作表
		XSSFSheet sheet = workbook.createSheet();
		int rownum=0;
		for (NewStudent stu : list) {
			XSSFRow row = sheet.createRow(rownum);
			row.createCell(0).setCellValue(stu.getId());
			row.createCell(1).setCellValue(stu.getName());
			row.createCell(2).setCellValue(stu.getScore());
			row.createCell(3).setCellValue(stu.getPhone());
			rownum++;	
		}
		//把工作博写入到服务器端
		System.out.println("创建的文件"+targetFile);
		workbook.write(new FileOutputStream(targetFile));
		//设置响应头
		response.setContentType("application/x-xls;charset=GBK");
		//设置浏览器下载提示
		response.setHeader("Content-Disposition", "attachment;filename=\"" + new String(fileName.getBytes(), "ISO8859-1") + "\"");
		//确定响应的文件的长度
		response.setContentLength((int)targetFile.length());
		//向响应文件流缓冲区写入文件
		byte[] buf = new byte[4096];
		BufferedOutputStream output = null;
		BufferedInputStream input = null;
		output = new BufferedOutputStream(response.getOutputStream());
		input = new BufferedInputStream(new FileInputStream(targetFile));
		//遍历文件
		int len = -1;
		while((len=input.read(buf))!=-1){
			output.write(buf,0,len);
		}
		output.flush();
		response.flushBuffer();
		if(input!=null){
			input.close();
		}if(output!=null){
			output.close();
		}

		
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	

}
