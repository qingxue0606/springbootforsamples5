package com.zhuozhengsoft.samples5.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Map;

@RestController
@RequestMapping(value = "/SendParameters/")
public class SendParametersController {
    private String dir = ResourceUtils.getURL("classpath:").getPath() + "static/doc/";

    public SendParametersController() throws FileNotFoundException {
    }

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面


        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);


        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法


        //打开Word文档
        poCtrl.webOpen("/doc/SendParameters/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SendParameters/Word");
        return mv;
    }


    @RequestMapping("save")
    public ModelAndView save(HttpServletRequest request, HttpServletResponse response,Map<String, Object> map) throws ClassNotFoundException, SQLException, FileNotFoundException {
        response.setCharacterEncoding("utf-8");
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SendParameters/" + fs.getFileName());
        int id = 0;
        String userName = "";
        int age = 0;
        String sex = "";

        //获取通过Url传递过来的值
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0)
            id = Integer.parseInt(request.getParameter("id").trim());

        //获取通过网页标签控件传递过来的参数值，注意fs.getFormField("参数名")方法中的参数名是值标签的“name”属性而不是Id

        //获取通过文本框<input type="text" />标签传递过来的值
        if (fs.getFormField("userName") != null
                && fs.getFormField("userName").trim().length() > 0) {
            userName = fs.getFormField("userName");
        }

        //获取通过隐藏域传递过来的值
        if (fs.getFormField("age") != null
                && fs.getFormField("age").trim().length() > 0) {
            age = Integer.parseInt(fs.getFormField("age"));
        }

        //获取通过<select>标签传递过来的值
        if (fs.getFormField("selSex") != null
                && fs.getFormField("selSex").trim().length() > 0) {
            sex = fs.getFormField("selSex");
        }
        Class.forName("org.sqlite.JDBC");//载入驱动程序类别
        String strUrl = "jdbc:sqlite:"
                +  ResourceUtils.getURL("classpath:").getPath() + "static/demodata/SendParameters.db";
        ;
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String strsql = "Update Users set UserName = '" + userName
                + "', age = " + age + ",sex = '" + sex + "' where id = "
                + id;
        stmt.executeUpdate(strsql);
        conn.close();
        fs.showPage(300, 200);
        fs.close();
        map.put("id", id);
        map.put("age", age);
        map.put("sex", sex);
        map.put("userName", userName);
        ModelAndView mv = new ModelAndView("SendParameters/Save");
        return mv;
    }

}
