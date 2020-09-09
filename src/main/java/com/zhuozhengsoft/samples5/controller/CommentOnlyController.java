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
import java.util.Map;

@RestController
@RequestMapping(value="/CommentOnly/")
public class CommentOnlyController {
    private String dir= ResourceUtils.getURL("classpath:").getPath()+"static\\doc\\";
    public CommentOnlyController() throws FileNotFoundException {
    }
    @RequestMapping(value="Word", method= RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String,Object> map){
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存","Save",1);
        poCtrl.addCustomToolButton("插入批注","newComment",0);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法


        //打开Word文档
        poCtrl.webOpen("/doc/CommentOnly/test.doc", OpenModeType.docCommentOnly,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("CommentOnly/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response){
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir+ "CommentOnly\\"+fs.getFileName());
        fs.close();
    }

}