package com.zhuozhengsoft.samples5.controller;

import java.io.FileNotFoundException;
import java.util.Map;
import java.util.Objects;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Controller;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.ModelAndView;

import com.zhuozhengsoft.pageoffice.*;

/**
 * @author Administrator
 *
 */
@RestController
public class DemoController {

	@Value("${popassword}") 
	private String poPassWord;

	
	@RequestMapping(value="/index", method=RequestMethod.GET)
	public ModelAndView showIndex(){
		ModelAndView mv = new ModelAndView("Index");
		return mv;
	}
	
	@RequestMapping(value="/word", method=RequestMethod.GET)
	public ModelAndView showWord(HttpServletRequest request, Map<String,Object> map){
		
		PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
		poCtrl.setServerPage("/poserver.zz");//设置服务页面
		poCtrl.addCustomToolButton("保存","Save",1);//添加自定义保存按钮
		poCtrl.addCustomToolButton("盖章","AddSeal",2);//添加自定义盖章按钮
		poCtrl.setSaveFilePage("/save");//设置处理文件保存的请求方法
		//打开word
		poCtrl.webOpen("d:\\test.doc",OpenModeType.docAdmin,"张三");
		map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
		
		ModelAndView mv = new ModelAndView("Word");
		return mv;
	}
	
	@RequestMapping("/save")
	public void saveFile(HttpServletRequest request, HttpServletResponse response){
		FileSaver fs = new FileSaver(request, response);
		fs.saveToFile("d:\\" + fs.getFileName());
		fs.close();
	}
	
	
	/**
	 * 添加PageOffice的服务器端授权程序Servlet（必须）
	 * @return
	 */
	@Bean
    public ServletRegistrationBean servletRegistrationBean() throws FileNotFoundException {
		com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();

		String poSysPath=ResourceUtils.getURL("classpath:").getPath()+"static\\lic";

		poserver.setSysPath(poSysPath);//设置PageOffice注册成功后,license.lic文件存放的目录
		ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
		srb.addUrlMappings("/poserver.zz");
		srb.addUrlMappings("/posetup.exe");
		srb.addUrlMappings("/pageoffice.js");
		srb.addUrlMappings("/jquery.min.js");
		srb.addUrlMappings("/pobstyle.css");
		srb.addUrlMappings("/sealsetup.exe");
        return srb;// 
    }
	
	/**
	 * 添加印章管理程序Servlet（可选）
	 * @return
	 */
	@Bean
    public ServletRegistrationBean servletRegistrationBean2() throws FileNotFoundException {
		com.zhuozhengsoft.pageoffice.poserver.AdminSeal adminSeal = new com.zhuozhengsoft.pageoffice.poserver.AdminSeal();
		adminSeal.setAdminPassword(poPassWord);//设置印章管理员admin的登录密码
		String poSysPath=ResourceUtils.getURL("classpath:").getPath()+"static\\lic";
		adminSeal.setSysPath(poSysPath);//设置印章数据库文件poseal.db存放的目录
		ServletRegistrationBean srb = new ServletRegistrationBean(adminSeal);
		srb.addUrlMappings("/adminseal.zz");
		srb.addUrlMappings("/sealimage.zz");
		srb.addUrlMappings("/loginseal.zz");
        return srb;// 
    }
}
